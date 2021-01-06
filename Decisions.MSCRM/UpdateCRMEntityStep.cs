using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Utilities;
using DecisionsFramework.Utilities.Data;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Decisions.MSCRM
{
    [Writable]
    public class UpdateCRMEntityStep : BaseCRMEntityStep, ISyncStep, IDataConsumer
    {
        public UpdateCRMEntityStep()
        {

        }
        public UpdateCRMEntityStep(string entityId)
        {
            EntityId = entityId;
        }

        public const string PATH_ERROR = "Error";
        public const string PATH_SUCCESS = "Success";
        private const string ENTITY_ID = "Entity Id";

        [WritableValue]
        private bool treatNullAsIgnore;

        [WritableValue]
        private bool treatEmptyStringAsNull;

        [PropertyClassification(new[] { "Settings" }, "Treat Null As Ignore", 0)]
        public bool TreatNullAsIgnore
        {
            get { return treatNullAsIgnore; }
            set { treatNullAsIgnore = value; }
        }

        [PropertyClassification(new[] { "Settings" }, "Treat Empty String As Null", 1)]
        public bool TreatEmptyStringAsNull
        {
            get { return treatEmptyStringAsNull; }
            set { treatEmptyStringAsNull = value; }
        }


        public override string StepName
        {
            get
            {
                return string.Format("Update {0} Entity", CRMEntity?.CRMEntityDisplayName);
            }
        }

        public override OutcomeScenarioData[] OutcomeScenarios
        {
            get
            {
                return new[]
                    {
                        new OutcomeScenarioData(PATH_SUCCESS, new DataDescription[0]),
                        new OutcomeScenarioData(PATH_ERROR, new DataDescription[] { new DataDescription(typeof(string), "Error Message") })
                    };
            }
        }

        public DataDescription[] InputData
        {
            get
            {
                List<DataDescription> inputData = new List<DataDescription>();
                inputData.Add(new DataDescription(typeof(string), ENTITY_ID));
                if (CRMEntity != null)
                {
                    Type type = GetMSCRMType();
                    if (type != null)
                    {
                        inputData.Add(new DataDescription(new DecisionsNativeType(type), CRMEntity.CRMEntityDisplayName));
                    }
                }
                return inputData.ToArray();
            }
        }
        public ResultData Run(StepStartData data)
        {
            try
            {
                if (CRMEntity == null) throw new Exception($"CRMEntity with id '{EntityId}' missing");
                if (data.Data.ContainsKey(ENTITY_ID))
                {
                    string entityId = data.Data[ENTITY_ID] as string;
                    CrmServiceClient conn = new CrmServiceClient(GetConnectionString());
                    IOrganizationService serviceProxy = conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;

                    Entity entity = new Entity(CRMEntity.CRMEntityName, new Guid(entityId));

                    if (data.Data.ContainsKey(CRMEntity.CRMEntityDisplayName))
                    {
                        Type type = GetMSCRMType();

                        object obj = data.Data[CRMEntity.CRMEntityDisplayName];
                        
                        Dictionary<string, object> objDict = (Dictionary<string, object>)type.GetProperty("Fields").GetValue(obj);

                        Type objType = obj.GetType();

                        CRMEntity.LogCrmEntityFields("UpdateStep");

                        foreach (var field in CRMEntity.CRMEntityFields)
                        {
                            if (field.IsValidForUpdate == true)
                            {
                                object fieldValue;

                                if (!objDict.TryGetValue(field.FieldName, out fieldValue))
                                {
                                    // If this field doesn't exist on the update obj, skip it.
                                    log.Debug($"Field '{field.FieldName}' is not being updated (no new value supplied).");
                                    continue;
                                }

                                if (TreatEmptyStringAsNull)
                                {
                                    string fieldValueStr = fieldValue as string;
                                    if(fieldValueStr != null && fieldValueStr.Length == 0)
                                    {
                                        log.Debug($"Empty string '{field.FieldName}' is being treated as null.");
                                        fieldValue = null;
                                    }
                                }

                                if (fieldValue == null)
                                {
                                    // If this option is set, don't add null values.
                                    if(!TreatNullAsIgnore) entity[field.FieldName] = null;
                                    if (TreatNullAsIgnore) log.Debug($"Null field '{field.FieldName}' is being ignored.");
                                    else log.Debug($"Null field '{field.FieldName}' updated normally.");
                                }
                                else
                                {
                                    if (field.AttributeType == AttributeTypeCode.Money.ToString())
                                    {
                                        decimal decimalValue = (decimal)fieldValue;
                                        entity[field.FieldName] = new Money(decimalValue);
                                        log.Debug($"Field '{field.FieldName}[Money/decimal]' updated normally.");
                                    }
                                    else if (field.AttributeType == AttributeTypeCode.Picklist.ToString())
                                    {
                                        SetPicklistValue(entity, field, fieldValue);
                                    }
                                    else if (field.AttributeType == AttributeTypeCode.Customer.ToString() || field.AttributeType == AttributeTypeCode.Lookup.ToString())
                                    {
                                        CRMLookUpTypeField lookUpFieldValue = fieldValue as CRMLookUpTypeField;
                                        if (lookUpFieldValue != null && string.IsNullOrEmpty(lookUpFieldValue.LookUpEntityName) == false && string.IsNullOrEmpty(lookUpFieldValue.Id) == false)
                                        {
                                            entity[field.FieldName] = new EntityReference(lookUpFieldValue.LookUpEntityName, new Guid(lookUpFieldValue.Id));
                                            log.Debug($"Field '{field.FieldName}[Lookup]' updated normally.");
                                        }
                                        else
                                        {
                                            log.Debug($"Field '{field.FieldName}' not updated: lookup field value not found.");
                                        }
                                    }
                                    else
                                    {
                                        entity[field.FieldName] = fieldValue;
                                        log.Debug($"Field '{field.FieldName}' updated normally.");
                                    }
                                }
                            }
                            else
                            {
                                log.Debug($"Skipping field '{field.FieldName}', not valid for update.");
                            }
                        }
                    }
                    serviceProxy.Update(entity);
                    return new ResultData(PATH_SUCCESS);
                }
                else
                {
                    return new ResultData(PATH_ERROR, new KeyValuePair<string, object>[] { new KeyValuePair<string, object>("Error Message", "Entity Id cannot be null.") });
                }
            }
            catch (Exception ex)
            {
                return new ResultData(PATH_ERROR, new KeyValuePair<string, object>[] { new KeyValuePair<string, object>("Error Message", ex.Message) });
            }
        }
    }
}
