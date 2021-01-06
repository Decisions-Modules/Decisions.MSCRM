using DecisionsFramework;
using DecisionsFramework.Data.ORMapper;
using DecisionsFramework.Design.Flow.Service;
using DecisionsFramework.ServiceLayer;
using DecisionsFramework.ServiceLayer.Actions;
using DecisionsFramework.ServiceLayer.Actions.Common;
using DecisionsFramework.ServiceLayer.Services.Folder;
using DecisionsFramework.ServiceLayer.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Decisions.MSCRM
{
   public class CRMFolderBehavior :  SystemFolderBehavior
    {
        public override BaseActionType[] GetFolderActions(Folder folder, BaseActionType[] proposedActions, EntityActionType[] types)
        {
            List<BaseActionType> list = new List<BaseActionType>(base.GetFolderActions(folder, proposedActions, types) ?? new BaseActionType[0]);
            list.Add(new EditObjectAction(typeof(CRMEntity), "Add CRM Entity", "", "", null,
                                                  new CRMEntity() { EntityFolderID = folder.FolderID },
                                                  new SetValueDelegate(AddCRMEntity))
            {
                ActionAddsType = typeof(CRMEntity),
                RefreshScope = ActionRefreshScope.OwningFolder
            });
            return list.ToArray();
        }


        private void AddCRMEntity(AbstractUserContext usercontext, object obj)
        {
            CRMEntity field = (CRMEntity)obj;

            new DynamicORM().Store(field);
        }

    }

    public class CRMStepsInitializer : IInitializable
    {
        const string CRM_LIST_FOLDER_ID = "CRM_ENTITY_FOLDER";
        const string CRM_PAGE_ID = "6f323bfd-0036-46aa-b5fe-7c4e23913419";

        public void Initialize()
        {
            ORM<Folder> orm = new ORM<Folder>();
            Folder folder = (Folder)orm.Fetch(typeof(Folder), CRM_LIST_FOLDER_ID);
            if (folder == null)
            {

                Log log = new Log("MSCRM Folder Behavior");
                log.Debug("Creating System Folder '" + CRM_LIST_FOLDER_ID + "'");
                folder = new Folder(CRM_LIST_FOLDER_ID, "MSCRM", Constants.INTEGRATIONS_FOLDER_ID);
                folder.FolderBehaviorType = typeof(CRMFolderBehavior).FullName;

                orm.Store(folder);
            }

            ORM<PageData> pageDataOrm = new ORM<PageData>();
            PageData pageData = pageDataOrm.Fetch(new WhereCondition[] {
                new FieldWhereCondition("configuration_storage_id", QueryMatchType.Equals, CRM_PAGE_ID),
                new FieldWhereCondition("entity_folder_id", QueryMatchType.Equals, CRM_LIST_FOLDER_ID)
            }).FirstOrDefault();

            if(pageData == null)
            {
                pageData = new PageData {
                    EntityFolderID = CRM_LIST_FOLDER_ID,
                    ConfigurationStorageID = CRM_PAGE_ID,
                    EntityName = "MSCRM Entities",
                    Order = -1
                };
                pageDataOrm.Store(pageData);
            }

            Folder typesFolder = orm.Fetch(CRMEntity.CRM_GENERATED_TYPES_FOLDER_ID);
            if(typesFolder == null)
            {
                Log log = new Log("MSCRM Folder Behavior");
                log.Debug("Creating System Folder '" + CRMEntity.CRM_GENERATED_TYPES_FOLDER_ID + "'");
                typesFolder = new Folder(CRMEntity.CRM_GENERATED_TYPES_FOLDER_ID, "MSCRM", Constants.DATA_STRUCTURES_FOLDER_ID);
                orm.Store(typesFolder);
            }

            FlowEditService.RegisterModuleBasedFlowStepFactory(new CRMStepsFactory());


        }
    }
}
