using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DecisionsFramework;
using DecisionsFramework.Utilities;

namespace Decisions.MSCRM
{
    [Writable]
    public class GetOptionFromValueStep : BaseCRMEntityStep, ISyncStep, IDataConsumer
    {
        public GetOptionFromValueStep() { }
        public GetOptionFromValueStep(string entityId)
        {
            EntityId = entityId;
        }

        const string INPUT_VALUE = "Option Value";
        public const string PATH_ERROR = "Error";
        public const string PATH_SUCCESS = "Success";
        const string OUTPUT_RESULT = "Option";

        [WritableValue]
        private string optionSetProperty;

        [PropertyClassification(0, "Option Set", "Settings")]
        [SelectStringEditor(nameof(AllOptionSetProperties))]
        public string OptionSetProperty
        {
            get { return optionSetProperty; }
            set
            {
                optionSetProperty = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(OutcomeScenarios));
            }
        }

        [PropertyHidden]
        public string[] AllOptionSetProperties
        {
            get
            {
                return CRMEntityFields?
                    .Where(field => field?.AttributeType == "Picklist")
                    .Select(field => field.FieldName).ToArray();
            }
        }

        public override string StepName => $"Get Option For {CRMEntity?.CRMEntityDisplayName ?? "Entity"}";

        public DataDescription[] InputData
        {
            get
            {
                List<DataDescription> inputData = new List<DataDescription>();
                inputData.Add(new DataDescription(typeof(int), INPUT_VALUE, false));
                return inputData.ToArray();
            }
        }

        public override OutcomeScenarioData[] OutcomeScenarios
        {
            get
            {
                return new[]
                    {
                        new OutcomeScenarioData(PATH_SUCCESS, new DataDescription[] { new DataDescription(GetOptionSetEnumType(), OUTPUT_RESULT, false) }),
                        new OutcomeScenarioData(PATH_ERROR, new DataDescription[] { new DataDescription(typeof(string), "Error Message") })
                    };
            }
        }

        internal Type GetOptionSetEnumType()
        {
            PropertyInfo prop = GetPropertyFromFieldName(GetMSCRMType(), optionSetProperty);
            if (prop == null)
                return null;
            Type enumType = prop.PropertyType.IsNullable() ?
                Nullable.GetUnderlyingType(prop.PropertyType)
                : prop.PropertyType;
            if (!enumType.IsEnum)
                return null;
            return enumType;
        }

        public ResultData Run(StepStartData data)
        {
            if (CRMEntity == null)
                throw new Exception($"CRMEntity with id '{EntityId}' missing");
            if (GetMSCRMType() == null)
                throw new Exception("Entity type not found");
            if (string.IsNullOrEmpty(optionSetProperty))
                throw new Exception("No option set chosen");

            Type enumType = GetOptionSetEnumType();
            if (enumType == null)
                throw new Exception("Couldn't find option set property " + optionSetProperty);
            if (!data.Data.ContainsKey(INPUT_VALUE) || !(data.Data[INPUT_VALUE] is int)) // Error path for anything that _could_ be handled in flow:
                return new ResultData(PATH_ERROR, new KeyValuePair<string, object>[] { new KeyValuePair<string, object>("Error Message", "No option value input was given") });

            CRMEntityField optionSetField = CRMEntityFields?.FirstOrDefault(field => field?.FieldName == optionSetProperty);
            if (optionSetField == null)
                return new ResultData(PATH_ERROR, new KeyValuePair<string, object>[] { new KeyValuePair<string, object>("Error Message", "No option set field found with name " + optionSetProperty) });

            // Start with the (CRM) option value:
            int optionValue = (int)data.Data[INPUT_VALUE];
            // Find the option with that value and get its name:
            string optionName = optionSetField.CRMOptionSet?
                .FirstOrDefault(option => option?.OptionValue != null && option.OptionValue.Value == optionValue)?.OptionName;
            if (string.IsNullOrEmpty(optionName))
                return new ResultData(PATH_ERROR, new KeyValuePair<string, object>[] { new KeyValuePair<string, object>("Error Message", $"No option found with value {optionValue}") });
            // The enum descriptions are the option names, so find the one that matches:
            string enumName = null;
            foreach (FieldInfo fieldInfo in enumType.GetFields(BindingFlags.Public | BindingFlags.Static))
            {
                DescriptionAttribute attr = fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false)
                    .FirstOrDefault() as DescriptionAttribute;

                if (attr?.Description == optionName)
                {
                    enumName = fieldInfo.Name;
                    break;
                }
            }
            if (enumName == null)
                throw new InvalidOperationException($"No enum constant found for '{optionName}' with value {optionValue}");

            object result = Enum.Parse(enumType, enumName);

            return new ResultData(PATH_SUCCESS, new KeyValuePair<string, object>[] { new KeyValuePair<string, object>(OUTPUT_RESULT, result) });
        }

        protected override IList<ValidationIssue> GetAdditionalValidationIssues()
        {
            List<ValidationIssue> issues = new List<ValidationIssue>();

            if (string.IsNullOrEmpty(optionSetProperty))
                issues.Add(new ValidationIssue(this, $"No option set property chosen", "", BreakLevel.Fatal));

            return issues.ToArray();
        }
    }
}
