using Decisions.MSCRM;
using DecisionsFramework.Data.ORMapper;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.DataStructure;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Flow.CoreSteps;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.ServiceLayer.Services.ContextData;
using DecisionsFramework.Utilities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Decisions.MSCRM
{
    [Writable]
    public class AddCRMEntityStep : BaseCRMEntityStep, ISyncStep, IDataConsumer
    {
        public AddCRMEntityStep()
        {

        }
        public AddCRMEntityStep(string entityId)
        {
            EntityId = entityId;
        }

        public const string PATH_ERROR = "Error";
        public const string PATH_SUCCESS = "Success";

        public override string StepName
        {
            get
            {
                return string.Format("Add {0} Entity", CRMEntity?.CRMEntityDisplayName);
            }
        }

        public override OutcomeScenarioData[] OutcomeScenarios
        {
            get
            {
                return new[]
                    {
                        new OutcomeScenarioData(PATH_SUCCESS, new DataDescription[] { new DataDescription(typeof(string), string.Format("Added {0} Id", CRMEntity?.CRMEntityDisplayName))}),
                        new OutcomeScenarioData(PATH_ERROR, new DataDescription[] { new DataDescription(typeof(string), "Error Message") })
                    };
            }
        }

        public DataDescription[] InputData
        {
            get
            {
                List<DataDescription> inputData = new List<DataDescription>();
                if (CRMEntity != null)
                {
                    Type type = GetMSCRMType();
                    if (type != null)
                        inputData.Add(new DataDescription(new DecisionsNativeType(type), CRMEntity.CRMEntityDisplayName));
                }
                return inputData.ToArray();
            }
        }

        public ResultData Run(StepStartData data)
        {
            try
            {
                if (CRMEntity == null) throw new Exception($"CRMEntity with id '{EntityId}' missing");
                CrmServiceClient conn = new CrmServiceClient(GetConnectionString());
                IOrganizationService serviceProxy = conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
                Entity entity = new Entity(CRMEntity.CRMEntityName);

                if (data.Data.ContainsKey(CRMEntity.CRMEntityDisplayName))
                {
                    Type type = GetMSCRMType();

                    object obj = data.Data[CRMEntity.CRMEntityDisplayName];

                    Dictionary<string, object> objDict = (Dictionary<string, object>)type.GetProperty("Fields").GetValue(obj);

                    foreach (var field in CRMEntity.CRMEntityFields)
                    {
                        if (field.IsValidForUpdate == true)
                        {
                            if (field.IsRequired && !objDict.ContainsKey(field.FieldName))
                            {
                                // Required fields must be present. If not, try to give a useful error message by listing all of the missing fields:
                                IEnumerable<string> missingFields = CRMEntity.CRMEntityFields
                                    .Where(f => f.IsValidForUpdate == true && f.IsRequired && !objDict.ContainsKey(f.FieldName))
                                    .Select(f => f.FieldName);
                                string allMissing = string.Join(", ", missingFields);
                                return new ResultData(PATH_ERROR, new KeyValuePair<string, object>[] {
                                    new KeyValuePair<string, object>("Error Message", string.Format("The following required fields are missing: {0}.", allMissing))
                                });
                            }

                            object fieldValue;

                            if (!objDict.TryGetValue(field.FieldName, out fieldValue))
                            {
                                // If this field doesn't exist on the update obj, skip it.
                                continue;
                            }

                            if (fieldValue == null)
                            {
                                entity[field.FieldName] = null;
                            }
                            else
                            {
                                if (field.AttributeType == AttributeTypeCode.Money.ToString())
                                {
                                    decimal decimalValue = (decimal)fieldValue;
                                    entity[field.FieldName] = new Money(decimalValue);
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
                                    }
                                }
                                else
                                {
                                    entity[field.FieldName] = fieldValue;
                                }
                            }
                        }
                    }
                }

                Guid guid = serviceProxy.Create(entity);
                return new ResultData(PATH_SUCCESS, new DataPair[] { new DataPair(string.Format("Added {0} Id", CRMEntity.CRMEntityDisplayName), guid.ToString()) });
            }
            catch (Exception ex)
            {
                return new ResultData(PATH_ERROR, new KeyValuePair<string, object>[] { new KeyValuePair<string, object>("Error Message", ex.Message) });
            }
        }
    }
}
