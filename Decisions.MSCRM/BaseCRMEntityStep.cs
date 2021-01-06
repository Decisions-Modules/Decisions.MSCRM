using DecisionsFramework;
using DecisionsFramework.Data.ORMapper;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Flow.CoreSteps;
using DecisionsFramework.Design.Flow.Interface;
using DecisionsFramework.Design.Flow.Service;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Utilities;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DescriptionAttribute = System.ComponentModel.DescriptionAttribute;

namespace Decisions.MSCRM
{
    [Writable]
    public abstract class BaseCRMEntityStep : BaseFlowAwareStep, IAddedToFlow, IValidationSource
    {
        protected static Log log = new Log("BaseCRMEntityStep");

        [WritableValue]
        private string entityId;

        [PropertyHidden]
        public string EntityId
        {
            get
            {
                return this.entityId;
            }
            set
            {
                this.entityId = value;
            }
        }

        [PropertyHidden]
        public CRMEntityField[] CRMEntityFields
        {
            get
            {
                if (CRMEntity != null)
                {
                    return CRMEntity.CRMEntityFields;
                }
                return null;
            }
        }

        private CRMEntity crmEntity;

        [PropertyHidden]
        public CRMEntity CRMEntity
        {
            get
            {
                if (crmEntity == null)
                {
                    ORM<CRMEntity> orm = new ORM<CRMEntity>();
                    crmEntity = orm.Fetch(EntityId);
                }
                return crmEntity;
            }
        }

        public abstract string StepName { get; }

        public Type GetMSCRMType()
        {
            return TypeUtilities.FindTypeByFullName(CRMEntity?.GetFullTypeName(), false);
        }

        protected string GetConnectionString() => CRMEntity.Connection.GetConnectionString();

        /// <summary>
        /// The property name will usually match the field name from MS CRM, but this method handles the cases where it was originally a reserved word.
        /// Returns null if type/name/property is missing.
        /// </summary>
        internal static PropertyInfo GetPropertyFromFieldName(Type crmEntityType, string fieldName)
        {
            if (crmEntityType == null || string.IsNullOrEmpty(fieldName))
                return null;

            PropertyInfo prop = crmEntityType.GetProperty(fieldName);
            if (prop == null && StringUtils.IsInvalidFieldName(fieldName))
            { // If not found, the original name might be a keyword, so do what GeneratorHelper.GetValidPropertyName does to find the actual property name:
                string newPropName = fieldName.First().ToString().ToUpper() + fieldName.Substring(1); // capitalize
                newPropName = StringUtils.StripInvalidCharactersFromLanguageName(newPropName); // remove illegal chars & add underscore if still a keyword
                prop = crmEntityType.GetProperty(newPropName);
            }
            return prop;
        }

        public object CreateObjectFromAttributes(Type type, CRMEntityField[] fields, AttributeCollection attributes)
        {
            var obj = Activator.CreateInstance(type);
            foreach (var attribute in attributes)
            {
                PropertyInfo pinfo = GetPropertyFromFieldName(type, attribute.Key);
                if (pinfo != null)
                {
                    object attributeValueObj;
                    if (attribute.Value is OptionSetValue)
                    {
                        int i = ((OptionSetValue)attribute.Value).Value;
                        Type enumType = pinfo.PropertyType.IsNullable() ?
                            Nullable.GetUnderlyingType(pinfo.PropertyType)
                            : pinfo.PropertyType;
                        if (!enumType.IsEnum)
                            throw new InvalidOperationException($"Type {enumType.FullName} is not an enum type");

                        // Find the option with this value (as defined in MSCRM), which was recorded in the CRMEntityFields:
                        CRMEntityField crmEntityField = fields?.FirstOrDefault(field => field.FieldName == attribute.Key);
                        CRMOptionsSet option = crmEntityField?.CRMOptionSet?.FirstOrDefault(o => o.OptionValue == i);
                        if (option == null)
                            throw new InvalidOperationException(
                                $"No option with value {i} found in enum {enumType.Name} for property {attribute.Key} on entity {type.Name}");

                        // Then find the enum constant where the description matches the option name:
                        string enumName = null;
                        foreach (FieldInfo fieldInfo in enumType.GetFields(BindingFlags.Public | BindingFlags.Static))
                        {
                            DescriptionAttribute attr = fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false)
                                .FirstOrDefault() as DescriptionAttribute;

                            if (attr != null && attr.Description == option.OptionName)
                            {
                                enumName = fieldInfo.Name;
                                break;
                            }
                        }
                        if (enumName == null)
                            throw new InvalidOperationException($"No enum constant found for '{option.OptionName}' with value {option.OptionValue}");

                        attributeValueObj = Enum.Parse(enumType, enumName);
                    }
                    else if (attribute.Value is Money)
                    {
                        attributeValueObj = ((Money)attribute.Value).Value;
                    }
                    else if (attribute.Value is EntityReference)
                    {
                        EntityReference lookFieldValue = attribute.Value as EntityReference;
                        attributeValueObj = new CRMLookUpTypeField() { LookUpEntityName = lookFieldValue.LogicalName, Id = lookFieldValue.Id.ToString() };
                    }
                    else
                    {
                        attributeValueObj = attribute.Value;
                    }

                    pinfo.SetValue(obj, attributeValueObj);
                }
            }
            return obj;
        }

        internal void SetPicklistValue(Entity entity, CRMEntityField field, object fieldValue)
        {
            if (field.CRMOptionSet.Length == 0) // Empty picklist - don't bother checking for matches
                return;
            int? optionValue = field.CRMOptionSet.FirstOrDefault(t => t.OptionName == fieldValue.ToString())?.OptionValue;
            if (optionValue != null)
            {
                entity[field.FieldName] = new OptionSetValue(optionValue.GetValueOrDefault());
                log.Debug($"Field '{field.FieldName}[OptionSet]' updated normally.");
            }
            else
            {
                log.Debug($"Value '{fieldValue.ToString()}' not found in optionset, checking enum type directly.");
                string enumTypeName = $"{CRMEntity.GetFullTypeName()}_{field.FieldName}";
                Type enumType = TypeUtilities.FindTypeByFullName(enumTypeName);
                if (enumType != null)
                {
                    bool enumFound = false;
                    try
                    {
                        if (Enum.IsDefined(enumType, fieldValue))
                        {
                            // Enum.Parse should return an object with the string representation instead of int:
                            object enumObj = Enum.Parse(enumType, fieldValue.ToString(), true);
                            enumFound = true;
                            // Then, find the Description of the enum, because this is what will actually match the dropdown list.
                            FieldInfo fieldInfo = enumObj.GetType().GetField(enumObj.ToString());
                            DescriptionAttribute[] attributes = (DescriptionAttribute[])fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);
                            string enumDisplayName = null;
                            if (attributes.Length > 0)
                            {
                                enumDisplayName = attributes[0].Description;
                            }
                            optionValue = field.CRMOptionSet.FirstOrDefault(t => t.OptionName == enumDisplayName)?.OptionValue;
                            if (optionValue != null)
                            {
                                entity[field.FieldName] = new OptionSetValue(optionValue.GetValueOrDefault());
                                log.Debug($"Field '{field.FieldName}[OptionSet]' updated normally.");
                            }
                            else
                            {
                                log.Debug($"Field '{field.FieldName}[OptionSet]' not updated: null OptionValue, or no option found matching name '{enumDisplayName}'.");
                            }
                        }
                    }
                    catch
                    {
                        // Enum not found...
                    }
                    finally
                    {
                        if (!enumFound && log.IsDebugEnabled)
                        {
                            log.Debug($"Field '{field.FieldName}' not updated: optionset value '{fieldValue.ToString()}' not found in optionset "
                            + $"({string.Join(", ", field.CRMOptionSet.Select(x => x.OptionName))})"
                            + $" or in enum type '{enumTypeName}'.");
                        }
                    }
                }
                else
                {
                    if (log.IsDebugEnabled)
                    {
                        log.Debug($"Field '{field.FieldName}' not updated: optionset value '{fieldValue.ToString()}' not found in optionset "
                        + $"({string.Join(", ", field.CRMOptionSet.Select(x => x.OptionName))})"
                        + " and enum type was not found.");
                    }
                }
            }
        }

        public void AddedToFlow()
        {
            if ((this.Flow != null) && (this.FlowStep != null))
                FlowEditService.SetDefaultStepName(this.Flow, this.FlowStep, StringUtils.SplitCamelCaseString(StepName) + " {0}");
        }

        public void RemovedFromFlow()
        {

        }

        protected virtual IList<ValidationIssue> GetAdditionalValidationIssues() => null;

        public ValidationIssue[] GetValidationIssues()
        {
            List<ValidationIssue> issues = new List<ValidationIssue>();

            if (CRMEntity == null)
                issues.Add(new ValidationIssue(this, $"CRM Entity with id '{EntityId}' is missing", "", BreakLevel.Fatal));

            var extraIssues = GetAdditionalValidationIssues();
            if (extraIssues != null)
                issues.AddRange(extraIssues);

            return issues.ToArray();
        }
    }
}
