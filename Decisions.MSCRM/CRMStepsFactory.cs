using Decisions.MSCRM;
using DecisionsFramework;
using DecisionsFramework.Data.ORMapper;
using DecisionsFramework.Design.Flow;
using DecisionsFramework.ServiceLayer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Decisions.MSCRM
{

    public class CRMStepsFactory : BaseFlowEntityFactory
    {
        const string PARENT_NODE = "MSCRM";
        public override string[] GetRootCategories(string flowId, string folderId)
        {
            // return root folders for each node
            return new string[] { "Data" };

        }

        public override FlowStepToolboxInformation[] GetFavoriteSteps(string flowId, string folderId)
        {
            return new FlowStepToolboxInformation[0];
        }
        public override string[] GetSubCategories(string[] nodes, string flowId, string folderId)
        {
            // return sub folders for each node
            /*if (nodes.Length == 1 && nodes[0] == "Data")
                return new string[] { PARENT_NODE };
            if (nodes != null && nodes.Length == 2 && nodes[1] == PARENT_NODE)
            {
                ORM<CRMEntity> orm = new ORM<CRMEntity>();
                CRMEntity[] crmEntities = orm.Fetch();
                return crmEntities.Select(t => t.CRMEntityDisplayName).ToArray();
            }
            return new string[0];*/
            if (nodes == null || nodes.Length == 0) return new string[0];
            if (nodes[0] == "Data")
            {
                if(nodes.Length == 1) return new string[] { PARENT_NODE };
                if(nodes[1] == PARENT_NODE)
                {
                    if (nodes.Length == 2)
                    {
                        return ModuleSettingsAccessor<CRMSettings>.Instance.Connections.Select(x => x.ConnectionName).ToArray();
                    }
                    else if (nodes.Length == 3)
                    {
                        CRMConnection connection = CRMConnection.GetCRMConnectionForName(nodes[2]);
                        if (connection == null) return new string[0];
                        ORM<CRMEntity> orm = new ORM<CRMEntity>();
                        CRMEntity[] crmEntities = orm.Fetch(new WhereCondition[]
                        {
                        new FieldWhereCondition("connection_id", QueryMatchType.Equals, connection.connectionId)
                        });
                        return crmEntities.Select(t => t.CRMEntityDisplayName).ToArray();
                    }
                }
            }
            return new string[0];

        }
        public override FlowStepToolboxInformation[] GetStepsInformation(string[] nodes, string flowId, string folderId)
        {
            // return step info
            if (nodes == null || nodes.Length != 4 || nodes[0] != "Data" || nodes[1] != PARENT_NODE)
                return new FlowStepToolboxInformation[0];

            List<FlowStepToolboxInformation> list = new List<FlowStepToolboxInformation>();
            CRMConnection connection = CRMConnection.GetCRMConnectionForName(nodes[2]);
            if(connection == null) return new FlowStepToolboxInformation[0];

            ORM<CRMEntity> orm = new ORM<CRMEntity>();
            CRMEntity[] crmEntities = orm.Fetch(new WhereCondition[] {
                new FieldWhereCondition("connection_id", QueryMatchType.Equals, connection.connectionId),
                new FieldWhereCondition("crm_entity_display_name", QueryMatchType.Equals, nodes[3])
            });

            foreach (CRMEntity entity in crmEntities)
            {
                list.Add(new FlowStepToolboxInformation("Get All Entities", nodes, string.Format("getAllEntities${0}", entity.entityId)));
                list.Add(new FlowStepToolboxInformation("Get Entity By Id", nodes, string.Format("getEntityById${0}", entity.entityId)));
                list.Add(new FlowStepToolboxInformation("Add Entity", nodes, string.Format("addCRMEntity${0}", entity.entityId)));
                list.Add(new FlowStepToolboxInformation("Update Entity", nodes, string.Format("updateCRMEntity${0}", entity.entityId)));
                list.Add(new FlowStepToolboxInformation("Delete Entity", nodes, string.Format("deleteEntity${0}", entity.entityId)));
                if (entity.CRMEntityFields?.Any(field => field?.AttributeType == "Picklist") == true)
                {
                    list.Add(new FlowStepToolboxInformation("Get Option For Entity", nodes, string.Format("getOptionForCRMEntity${0}", entity.entityId)));
                }
            }
            return list.ToArray();
        }

        public override IFlowEntity CreateStep(string[] nodes, string stepId, StepCreationInfo additionalInfo)
        {
            string[] parts = (stepId ?? string.Empty).Split('$');
            if (parts.Length < 2)
                return null;

            string crmEntityId = parts[1];
            // create steps
            if (stepId.StartsWith("addCRMEntity"))
            {
                return new AddCRMEntityStep(crmEntityId);
            }
            if (stepId.StartsWith("updateCRMEntity"))
            {
                return new UpdateCRMEntityStep(crmEntityId);
            }
            if (stepId.StartsWith("getAllEntities"))
            {
                return new GetAllCRMEntitiesStep(crmEntityId);
            }
            if (stepId.StartsWith("deleteEntity"))
            {
                return new DeleteCRMEntityStep(crmEntityId);
            }
            if (stepId.StartsWith("getEntityById"))
            {
                return new GetCRMEntityByIdStep(crmEntityId);
            }
            if (stepId.StartsWith("getOptionForCRMEntity"))
            {
                return new GetOptionFromValueStep(crmEntityId);
            }
            return null;
        }

        public override FlowStepToolboxInformation[] SearchSteps(string flowId, string folderId, string searchString, int maxRecords)
        {
            // return null;
            return new FlowStepToolboxInformation[0];
        }
    }
}
