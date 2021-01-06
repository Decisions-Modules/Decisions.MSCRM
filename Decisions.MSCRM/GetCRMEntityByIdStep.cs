using DecisionsFramework.Design.Flow;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.ServiceLayer.Services.ContextData;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Decisions.MSCRM
{
    public class GetCRMEntityByIdStep : BaseCRMEntityStep, ISyncStep, IDataConsumer
    {
        public GetCRMEntityByIdStep()
        {

        }
        public GetCRMEntityByIdStep(string entityId)
        {
            EntityId = entityId;
        }

        public const string PATH_ERROR = "Error";
        public const string PATH_SUCCESS = "Success";
        const string ENTITY_ID = "Entity Id";
        public override string StepName
        {
            get
            {
                return string.Format("Get {0} Entity By Id", CRMEntity?.CRMEntityDisplayName);
            }
        }

        public override OutcomeScenarioData[] OutcomeScenarios
        {
            get
            {
                Type type = GetMSCRMType();
                if (type == null)
                    return new OutcomeScenarioData[0];

                return new[]
                    {
                        new OutcomeScenarioData(PATH_SUCCESS, new DataDescription[] { new DataDescription(type, string.Format("{0} Entity", CRMEntity?.CRMEntityDisplayName), false)}),
                        new OutcomeScenarioData(PATH_ERROR, new DataDescription[] { new DataDescription(typeof(string), "Error Message") })
                    };
            }
        }

        public DataDescription[] InputData
        {
            get
            {
                return new DataDescription[]
                {
                    new DataDescription(new DecisionsNativeType(typeof(string)),ENTITY_ID,false,false,false)
                };
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

                    Entity entity = serviceProxy.Retrieve(CRMEntity.CRMEntityName, new Guid(entityId), new Microsoft.Xrm.Sdk.Query.ColumnSet(true));

                    if(entity != null)
                    {
                        Type type = GetMSCRMType();

                        var obj = CreateObjectFromAttributes(type, CRMEntity.CRMEntityFields, entity.Attributes);

                        return new ResultData(PATH_SUCCESS, new DataPair[] { new DataPair(string.Format("{0} Entity", CRMEntity.CRMEntityDisplayName), obj) });
                    }
                    else
                    {
                        return new ResultData(PATH_ERROR, new KeyValuePair<string, object>[] { new KeyValuePair<string, object>("Error Message", "No entity with that ID could be found.") });
                    }
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
