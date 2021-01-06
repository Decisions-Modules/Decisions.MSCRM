using DecisionsFramework;
using DecisionsFramework.Data.ORMapper;
using DecisionsFramework.Design.ConfigurationStorage.Attributes;
using DecisionsFramework.Design.DataStructure;
using DecisionsFramework.Design.Flow.Mapping;
using DecisionsFramework.Design.Properties;
using DecisionsFramework.Design.Properties.Attributes;
using DecisionsFramework.ServiceLayer;
using DecisionsFramework.ServiceLayer.Actions;
using DecisionsFramework.ServiceLayer.Actions.Common;
using DecisionsFramework.ServiceLayer.Services.Folder;
using DecisionsFramework.ServiceLayer.Utilities;
using DecisionsFramework.Utilities.CodeGeneration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Decisions.MSCRM
{

    [Writable]
    [ORMEntity]
    [ValidationRules]
    [Exportable]
    public class CRMEntity : AbstractFolderEntity, INotifyPropertyChanged, IValidationSource
    {
        private static Log log = new Log("CRMEntity");

        public const string CRM_ENTITY_NAMESPACE = "Decisions.MSCRM";
        public const string CRM_GENERATED_TYPES_FOLDER_ID = "MSCRM_GENERATED_TYPES_FOLDER";
  
        #region Fields
        [WritableValue]
        [ORMPrimaryKeyField]
        [PropertyHidden]
        public string entityId;

        [WritableValue]
        [ORMField]
        private string crmEntityName;

        [WritableValue]
        [ORMField]
        private string crmEntityDisplayName;

        [WritableValue]
        [ORMField]
        private string connectionId;

        [WritableValue]
        [ORMField]
        private string connectionName;

        [WritableValue]
        [ORMField]
        private string organisationUrl;

        [WritableValue]
        [ORMField]
        private string domain;

        [WritableValue]
        [ORMField]
        private string userName;

        [WritableValue]
        [ORMField]
        private string password;

        [WritableValue]
        [ORMField]
        private bool overrideConnectionString;

        [WritableValue]
        [ORMField]
        private string connectionString;

        private CRMConnection connection;

        [WritableValue]
        [ORMField(typeof(ORMXmlSerializedFieldConverter))]
        private CRMEntityField[] Fields;

        #endregion
        #region Properties

        [PropertyClassification(1, "CRM Entity Name", "[Settings]")]
        [RequiredProperty("CRM Entity Name is required.")]
        [SelectStringEditor("AllEntityNames")]
        public override string EntityName
        {
            get
            {
                return base.EntityName;
            }

            set
            {
                base.EntityName = value;
                OnPropertyChanged(nameof(EntityName));
            }
        }

        [BooleanPropertyHidden("ShowDescription", false)]
        [ExcludeInDescription]
        [RichTextInputEditor]
        [PropertyClassification(2, "Description", "[Settings]")]
        public override string EntityDescription
        {
            get
            {
                return base.EntityDescription;
            }

            set
            {
                base.EntityDescription = value;
            }
        }

        [PropertyClassification("Connection Name", 0)]
        [RequiredProperty("Connection Name is required.")]
        [SelectStringEditor("AllConnectionNames")]
        public string ConnectionName
        {
            get
            {
                return connectionName;
            }
            set
            {
                if (value != connectionName)
                {
                    connectionName = value;
                    Connection = null; // Refresh 'Connection'. Entity name lists will also be refreshed.
                    allEntityNames = null; // Get rid of these anyway, in case the connection was already null.
                    allEntityLogicalNames = null;
                    allEntityDisplayNames = null;
                    OnPropertyChanged(nameof(ConnectionName));
                    OnPropertyChanged(nameof(AllEntityNames));
                    OnPropertyChanged(nameof(EntityName));
                }
            }
        }

        [PropertyHidden]
        public string CRMEntityName
        {
            get
            {
                return crmEntityName;
            }
            set
            {
                crmEntityName = value;
            }
        }

        [PropertyHidden]
        public string CRMEntityDisplayName
        {
            get
            {
                return crmEntityDisplayName;
            }
            set
            {
                crmEntityDisplayName = value;
            }
        }

        [PropertyHidden]
        public string[] AllConnectionNames
        {
            get
            {
                CRMConnection[] connections = ModuleSettingsAccessor<CRMSettings>.Instance.Connections;
                if(connections == null || connections.Length == 0)
                {
                    return new[] { "Configure connections in /System/Settings" };
                }
                return connections.Select(x => x.ConnectionName).ToArray();
            }
            set { }
        }

        // For caching only - not saved to db.
        private string[] allEntityNames; // (combined)
        private string[] allEntityLogicalNames;
        private string[] allEntityDisplayNames;

        [PropertyHidden]
        public string[] AllEntityNames
        {
            get
            {
                if (allEntityNames != null) return allEntityNames;

                if (Connection == null) return new[] { "Invalid connection name" };

                string[] logicalNames = Connection.AllEntityNames;
                string[] displayNames = Connection.AllEntityDisplayNames;

                if (logicalNames == null || displayNames == null || logicalNames.Length != displayNames.Length) return new[] { "Error retrieving entity names" };

                allEntityLogicalNames = logicalNames;
                allEntityDisplayNames = displayNames;

                allEntityNames = new string[allEntityLogicalNames.Length];
                for (int i = 0; i < allEntityNames.Length; ++i)
                {
                    allEntityNames[i] = string.Format("{0} ({1})", allEntityDisplayNames[i], allEntityLogicalNames[i]);
                }

                return allEntityNames;
            }
            set { }
        }

        [PropertyHidden]
        public CRMConnection Connection
        {
            get
            {
                if (connection == null)
                {
                    connection = CRMConnection.GetCRMConnectionForName(connectionName);
                }
                return connection;
            }
            set
            {
                if(connection != value)
                {
                    connection = value;
                    allEntityNames = null;
                    allEntityLogicalNames = null;
                    allEntityDisplayNames = null;
                }
            }
        }

        [PropertyHidden]
        public CRMEntityField[] CRMEntityFields
        {
            get
            {
                return Fields;
            }
            set
            {
                LogCrmEntityFields("BeforeSet");

                if(value == null || value.Length == 0)
                {
                    log.Debug("Null or empty assigned to CRMEntityFields. Stack trace: " + Environment.StackTrace);
                }
                /*string logFieldsNull = Fields == null ? "null" : Fields.Length.ToString();
                if (value == null) log.Debug($"Null assigned to CRMEntityFields, previous length was {logFieldsNull}.");
                else if (value.Length == 0) log.Debug("Empty array assigned to CRMEntityFields, previous length was {logFieldsNull}.");*/

                Fields = value;
                LogCrmEntityFields("AfterSet");
            }
        }

        #endregion

        internal void LogCrmEntityFields(string extraDescription = "")
        {
            List<string> lines = new List<string> { $"LogCrmEntityFields ({CRMEntityName})[{extraDescription}]:" };
            try
            {

                if (CRMEntityFields == null) lines.Add("(null)");
                else if (CRMEntityFields.Length == 0) lines.Add("(empty)");
                else
                {
                    int i = 0;
                    foreach (CRMEntityField field in CRMEntityFields)
                    {

                        StringBuilder sb = new StringBuilder();
                        sb.Append($"{(i++).ToString().PadLeft(5)} : {field.FieldName?.PadRight(50)}({field.DisplayName?.PadRight(50)})[{field.AttributeType?.PadRight(50)}/{field.FieldType.SilverlightAssemblyQualifiedName?.PadRight(150)}], ");
                        sb.Append($"ValidForUpdate:{field.IsValidForUpdate.ToString().PadRight(5)}, Required:{field.IsRequired.ToString().PadRight(5)}");
                        if (field.CRMOptionSet != null)
                        {
                            sb.Append(", Optionset: ");
                            sb.Append(string.Join(", ", field.CRMOptionSet.Select(x => $"({x.OptionName}:{x.OptionValue}")));
                        }

                        lines.Add(sb.ToString());
                    }
                }
            }
            catch(Exception e)
            {
                // Not letting logging throw exceptions...
                lines.Add("(caught exception during logging: " + e.Message);
            }
            finally
            {
                log.Debug(string.Join("    \r\n", lines));
            }
        }

        public override void AfterRead()
        {
            LogCrmEntityFields("AfterRead");

            base.AfterRead();
        }

        public override void BeforeSave()
        {
            CRMConnection specifiedConnection = connection;
            if (specifiedConnection == null)
            {
                // If 'connection' is null, this might be an import.
                log.Debug("sConnection is null, fetching by ID");
                specifiedConnection = CRMConnection.GetCRMConnectionById(connectionId);
                if (specifiedConnection != null) Connection = specifiedConnection;
            }
            // If a connection exists at this point, make sure this entity name doesn't already exist for this connection:
            if(specifiedConnection != null)
            {
                SetNamesFromSelectedName();
                log.Debug($"Checking whether entity already exists with connection_id '{specifiedConnection.connectionId}' and crm_entity_name '{CRMEntityName}'.");
                ORM<CRMEntity> orm = new ORM<CRMEntity>();
                var conditions = new List<WhereCondition>
                {
                    new FieldWhereCondition("connection_id", QueryMatchType.Equals, specifiedConnection.connectionId),
                    new FieldWhereCondition("crm_entity_name", QueryMatchType.Equals, this.CRMEntityName)
                };
                if (!string.IsNullOrWhiteSpace(this.entityId))
                {
                    // (if ID is the same, this is an edit)
                    conditions.Add(new FieldWhereCondition("entity_id", QueryMatchType.DoesNotEqual, this.entityId));
                }
                CRMEntity otherEntity = orm.Fetch(conditions.ToArray()).FirstOrDefault();
                log.Debug($"entity: {otherEntity?.CRMEntityDisplayName ?? "(null)"}");
                if (otherEntity != null) throw new InvalidOperationException("This entity already exists for this connection.");
            }

            if(specifiedConnection == null)
            {
                // If the ID is missing, this might be an import. Check for a matching name:
                log.Debug("sConnection is null, fetching by name");
                specifiedConnection = CRMConnection.GetCRMConnectionForName(connectionName);
            }
            if (specifiedConnection == null)
            {
                // If no connection was found by ID or by name, create one:
                log.Debug("sConnection is null, creating");
                specifiedConnection = new CRMConnection()
                {
                    ConnectionName = connectionName,
                    OrganisationUrl = organisationUrl,
                    Domain = domain,
                    UserName = userName,
                    Password = password,
                    OverrideConnectionString = overrideConnectionString,
                    ConnectionString = connectionString
                };
                // Add new connection to settings:
                CRMConnection[] oldConnections = ModuleSettingsAccessor<CRMSettings>.Instance.Connections;
                ModuleSettingsAccessor<CRMSettings>.Instance.Connections = oldConnections.Concat(new[] { specifiedConnection }).ToArray();
                log.Debug($"about to save new connections...");
                ModuleSettingsAccessor<CRMSettings>.SaveSettings();
                specifiedConnection = CRMConnection.GetCRMConnectionForName(connectionName);
                if (specifiedConnection == null) throw new EntityNotFoundException("CRMConnection was not created successfully.");
                log.Debug("new connections saved.");
            }
            if (specifiedConnection != null)
            {
                // Update our data to match the connection's data:
                log.Debug("sConnection exists, updating data to match it...");
                Connection = specifiedConnection;
                connectionId = specifiedConnection.connectionId;
                connectionName = specifiedConnection.ConnectionName;
                organisationUrl = specifiedConnection.OrganisationUrl;
                domain = specifiedConnection.Domain;
                userName = specifiedConnection.UserName;
                password = specifiedConnection.Password;
                overrideConnectionString = specifiedConnection.OverrideConnectionString;
                connectionString = specifiedConnection.ConnectionString;
            }

            SetNamesFromSelectedName();

            base.BeforeSave();

            AddOrUpdateCRMEntity();
        }

        private void SetNamesFromSelectedName()
        {
            int idx = Array.IndexOf(AllEntityNames, EntityName);
            if (idx == -1) throw new InvalidOperationException("Cannot create CRMEntity: Entity name must match name in list.");
            CRMEntityName = allEntityLogicalNames[idx];
            CRMEntityDisplayName = allEntityDisplayNames[idx];
        }

        public override void BeforeDelete()
        {
            base.BeforeDelete();
            ORM<SimpleFlowStructure> orm = new ORM<SimpleFlowStructure>();
            orm.Delete(new WhereCondition[]
            {
                new FieldWhereCondition("data_type_name", QueryMatchType.Equals, CRMEntityName),
                new FieldWhereCondition("data_type_name_space", QueryMatchType.Equals, GetCrmEntityNamespace())
            });

            ORM<Folder> folderOrm = new ORM<Folder>();
            folderOrm.Delete(new WhereCondition[] {
                new FieldWhereCondition("folder_id", QueryMatchType.Equals, GetCrmEntityFolderId())
            });
        }


        private void AddOrUpdateCRMEntity()
        {
            IOrganizationService serviceProxy = GetCRMClientServiceProxy();
            if (serviceProxy == null)
                throw new Exception("Unable to found CRM client service proxy");
            try
            {
                ORM<SimpleFlowStructure> orm = new ORM<SimpleFlowStructure>();
                //Check current type of entity present in data base or not
                SimpleFlowStructure simpleFlowStructure = orm.Fetch(new WhereCondition[]
                {
                        new FieldWhereCondition("data_type_name", QueryMatchType.Equals, CRMEntityName),
                        new FieldWhereCondition("data_type_name_space", QueryMatchType.Equals, GetCrmEntityNamespace())
                }).FirstOrDefault();
                
                if (simpleFlowStructure == null)
                {
                    EnsureCrmEntityFolderExists();
                    simpleFlowStructure = new SimpleFlowStructure();
                    simpleFlowStructure.DataTypeNameSpace = GetCrmEntityNamespace();
                    simpleFlowStructure.EntityFolderID = GetCrmEntityFolderId();
                    simpleFlowStructure.StorageOption = StorageOption.NotDatabaseStored;
                    simpleFlowStructure.SuperClass = null;
                    simpleFlowStructure.TemplateForType = @"DecisionsFramework.Utilities.CodeGeneration.Templates.StringMappedDataStructure.vm";
                }
                simpleFlowStructure.DataTypeName = CRMEntityName;

                AddOrUpdateCRMEntityWithDataStructure(serviceProxy, CRMEntityName, simpleFlowStructure);
            }
            catch (Exception ex)
            {
                log.Error(ex);
                throw ex;
            }
        }

        internal string GetFullTypeName()
        {
            return GetCrmEntityNamespace() + "." + CRMEntityName;
        }

        private string GetCrmEntityNamespace()
        {
            if (Connection == null) throw new InvalidOperationException("No connection found for this entity.");
            return CRM_ENTITY_NAMESPACE + ".crm_" + Connection.connectionId.Replace("-", "");
        }

        private string GetCrmEntityFolderId()
        {
            string connectionFolderId = Connection?.GetCrmConnectionFolderId();
            if (connectionFolderId == null) throw new InvalidOperationException("No connection folder found for this entity.");
            return connectionFolderId + CRMEntityName;
        }

        private void EnsureCrmEntityFolderExists()
        {
            if (Connection == null) throw new InvalidOperationException("Could not find connection.");

            Connection.EnsureCrmConnectionFolderExists();

            string connectionFolderId = Connection.GetCrmConnectionFolderId();

            ORM<Folder> folderORM = new ORM<Folder>();
            string crmEntityFolderId = GetCrmEntityFolderId();
            Folder crmEntityFolder = folderORM.Fetch(crmEntityFolderId);
            if (crmEntityFolder == null)
            {
                crmEntityFolder = new Folder(crmEntityFolderId, CRMEntityDisplayName, connectionFolderId);
                folderORM.Store(crmEntityFolder);
            }
        }

        private IOrganizationService GetCRMClientServiceProxy()
        {
            try
            {
                CrmServiceClient conn = new CrmServiceClient(Connection.GetConnectionString());
                if (conn == null)
                    throw new Exception("Unable to connect CRM client server");
                log.Info("Successfully connected to CRM client service");
                return conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            }
            catch (System.IO.FileNotFoundException)
            {
                // FileNotFound is sometimes thrown from within Microsoft.Xrm.Tooling.Connector, probably when the connection info is invalid.
                throw new InvalidOperationException("Connection could not be made - check connection info or retry.");
            }
        }

        private RetrieveEntityResponse GetEntityInformationFromCRMClient(IOrganizationService serviceProxy, string crmEntityName)
        {
            try
            {
                log.Info("Fetching CRM entity from server");
                RetrieveEntityRequest entitiesRequest = new RetrieveEntityRequest()
                {
                    EntityFilters = EntityFilters.Attributes,
                    LogicalName = crmEntityName,
                };
                RetrieveEntityResponse entity = (RetrieveEntityResponse)serviceProxy.Execute(entitiesRequest);
                return entity;
            }
            catch (Exception ex)
            {
                log.Error(ex, string.Format("Error occur while fetch CRM entity of name {0} from server", crmEntityName));
                throw ex;
            }
        }

        private void GetEntityFieldsFromCRMEntityAttributes(List<AttributeMetadata> entityAttributes, string crmEntityName, out List<CRMEntityField> entityFields, out List<DefinedDataTypeDataMember> memberList)
        {
            entityFields = new List<CRMEntityField>();
            memberList = new List<DefinedDataTypeDataMember>();
            if (entityAttributes.Count() > 0)
            {
                if (log.IsDebugEnabled)
                {
                    log.Debug(string.Join(", ", entityAttributes.Select(x => $"{x.LogicalName}({x.DisplayName})[{x.AttributeType}, ")));
                }

                foreach (AttributeMetadata attribute in entityAttributes)
                {
                    Type type = GetTypeFromAttributeType(attribute.AttributeType);
                    if (type != null)
                    {
                        string displayName;
                        if (attribute.DisplayName.LocalizedLabels != null && attribute.DisplayName.LocalizedLabels.Count > 0)
                        {
                            displayName = attribute.DisplayName.LocalizedLabels[0].Label;
                        }
                        else displayName = attribute.LogicalName;
                        bool? validForUpdate = attribute.IsValidForUpdate;
                        DefinedDataTypeDataMember member = new DefinedDataTypeDataMember();
                        member.RelationshipName = attribute.LogicalName;
                        member.DisplayLabelName = CleanDisplayName(displayName);
                        member.IsOverrideDisplayInformation = true;
                        member.RelatedToDataType = type.FullName;
                        member.AllowNull = true;
 
                        if (attribute.AttributeType == AttributeTypeCode.Uniqueidentifier || attribute.IsPrimaryId == true)
                        {
                            member.CustomAttributes = new CustomAttributeDefinition[] { new CustomAttributeDefinition("PropertyHidden", new string[0]) };
                            validForUpdate = false;
                        }
                        CRMEntityField entityField = new CRMEntityField()
                        {
                            DisplayName = CleanDisplayName(displayName),
                            FieldName = attribute.LogicalName,
                            FieldType = new DecisionsNativeType(type),
                            IsValidForUpdate = validForUpdate,
                            AttributeType = attribute.AttributeType.ToString()
                        };

                        if (attribute.AttributeType == AttributeTypeCode.Picklist)
                        {
                            ORM<EnumDataType> orm = new ORM<EnumDataType>();

                            //Check this type of enum data is present in database or not
                            // and either create it or update it
                            string currentOptionSetDataTypeName = string.Format("{0}_{1}", crmEntityName, attribute.LogicalName);

                            EnumDataType optionsSet = orm.Fetch(new WhereCondition[]
                            {
                                new FieldWhereCondition("data_type_name", QueryMatchType.Equals, currentOptionSetDataTypeName),
                                new FieldWhereCondition("data_type_name_space", QueryMatchType.Equals, GetCrmEntityNamespace())
                            }).FirstOrDefault();

                            if (optionsSet == null)
                            {
                                log.Debug($"No enum type with name '{currentOptionSetDataTypeName}' and namespace '{GetCrmEntityNamespace()}' found, trying to create one:");
                                EnsureCrmEntityFolderExists();
                                optionsSet = new EnumDataType();
                                optionsSet.DataTypeName = currentOptionSetDataTypeName;
                                optionsSet.DataTypeNameSpace = GetCrmEntityNamespace();
                                optionsSet.EntityFolderID = GetCrmEntityFolderId();
                                optionsSet.EnumValues = CleanEnumValues(((PicklistAttributeMetadata)attribute).OptionSet.Options.Select(t => t.Label.LocalizedLabels[0].Label).ToArray());
                                if (optionsSet.EnumValues.Length == 0)
                                    optionsSet.EnumValues = new string[] { "NoValue" };
                                orm.Store(optionsSet);
                                log.Debug($"new enum type created for '{currentOptionSetDataTypeName}'");
                            }
                            else
                            {
                                log.Debug($"Enum type '{GetCrmEntityNamespace()}.{currentOptionSetDataTypeName}' already exists; updating its enum value list.");

                                string[] enumValues = CleanEnumValues(((PicklistAttributeMetadata)attribute).OptionSet.Options.Select(t => t.Label.LocalizedLabels[0].Label).ToArray());

                                if (log.IsDebugEnabled)
                                {
                                    log.Debug(string.Join(", ", ((PicklistAttributeMetadata)attribute).OptionSet.Options.Select(t => $"'{t.Label.LocalizedLabels[0].Label}':'{t.Value}'")));
                                }

                                if (enumValues.Length == 0)
                                    enumValues = new string[] { "NoValue" };

                                DataStructureService.Instance.ChangeEnumValues(UserContextHolder.GetRootUserContext(), GetCrmEntityNamespace(), currentOptionSetDataTypeName, enumValues);

                                log.Debug("Enum value list updated.");
                            }

                            string generatedEnumFullTypeName = string.Format("{0}.{1}", GetCrmEntityNamespace(), currentOptionSetDataTypeName);

                            member.RelatedToDataType = generatedEnumFullTypeName;
                            List<CRMOptionsSet> options = new List<CRMOptionsSet>();
                            foreach (var option in ((PicklistAttributeMetadata)attribute).OptionSet.Options)
                            {
                                options.Add(new CRMOptionsSet() { OptionName = option.Label.LocalizedLabels[0].Label, OptionValue = option.Value });
                            }
                            entityField.CRMOptionSet = options.ToArray();
                        }

                        if (attribute.RequiredLevel.Value == AttributeRequiredLevel.ApplicationRequired || attribute.RequiredLevel.Value == AttributeRequiredLevel.SystemRequired)
                        {
                            entityField.IsRequired = true;
                            member.Required = true;
                        }

                        entityFields.Add(entityField);
                        memberList.Add(member);
                    }
                }
            }
        }

        private string CleanDisplayName(string displayName)
        {
            return displayName.Replace("\\", string.Empty);
        }

        private string[] CleanEnumValues(string[] v)
        { // Escape any quotes, since these strings are inserted into Description attributes:
            return v?.Select(s => s?.Replace("\"", "\\\"")).ToArray();
        }

        private void AddOrUpdateCRMEntityWithDataStructure(IOrganizationService serviceProxy, string crmEntityName, SimpleFlowStructure simpleFlowStructure)
        {
            RetrieveEntityResponse crmEntity = GetEntityInformationFromCRMClient(serviceProxy, crmEntityName);
            if (crmEntity != null)
            {
                log.Info("CRM entity fetched successfully from server");
                List<AttributeMetadata> entityAttributes = crmEntity.EntityMetadata.Attributes.Where(t => t.IsValidForCreate == true).ToList();
                List<CRMEntityField> entityFields;
                List<DefinedDataTypeDataMember> memberList;

                GetEntityFieldsFromCRMEntityAttributes(entityAttributes, crmEntityName, out entityFields, out memberList);

                simpleFlowStructure.Children = memberList.ToArray();
                CRMEntityFields = entityFields.ToArray();
                string displayName;
                if (crmEntity.EntityMetadata.DisplayName.LocalizedLabels != null && crmEntity.EntityMetadata.DisplayName.LocalizedLabels.Count > 0)
                {
                    displayName = crmEntity.EntityMetadata.DisplayName.LocalizedLabels[0].Label;
                }
                else displayName = crmEntity.EntityMetadata.LogicalName;
                CRMEntityDisplayName = displayName;
                DataStructureService.Instance.AddDataStructure(UserContextHolder.GetRootUserContext(), simpleFlowStructure);
                log.Info(string.Format("Entity with name {0} added successfully to the database", crmEntityName));
            }
        }

        private void RegenerateDataStructureOfSelectedEntity(AbstractUserContext userContext, string entityId)
        {
            try
            {
                ORM<CRMEntity> crmEntityORM = new ORM<CRMEntity>();
                CRMEntity crmEntity = crmEntityORM.Fetch(entityId);
                if (crmEntity != null)
                {
                    log.Info(string.Format("started regenerating data structure of {0} entity", crmEntity.CRMEntityName));
                    ORM<SimpleFlowStructure> simpleFlowStructureORM = new ORM<SimpleFlowStructure>();
                    SimpleFlowStructure simpleFlowStructure = simpleFlowStructureORM.Fetch(new WhereCondition[]
                    {
                             new FieldWhereCondition("data_type_name", QueryMatchType.Equals, crmEntity.CRMEntityName),
                             new FieldWhereCondition("data_type_name_space", QueryMatchType.Equals, crmEntity.GetCrmEntityNamespace())
                    }).FirstOrDefault();
                    if (simpleFlowStructure != null)
                    {
                        IOrganizationService serviceProxy = GetCRMClientServiceProxy();
                        if (serviceProxy == null)
                            throw new BusinessRuleException("Unable to found CRM service client proxy");

                        simpleFlowStructure.Children = null;
                        AddOrUpdateCRMEntityWithDataStructure(serviceProxy, crmEntity.CRMEntityName, simpleFlowStructure);
                        // Make sure the updated CRMEntityFields are stored:
                        crmEntityORM.Store(this, true, false);
                    }
                    log.Info(string.Format("completed regenerating data structure of {0} entity", crmEntity.CRMEntityName));
                }
            }
            catch (Exception ex)
            {
                log.Error(ex);
                throw ex;
            }

        }

        public override BaseActionType[] GetActions(AbstractUserContext userContext, EntityActionType[] types)
        {
            List<BaseActionType> actions = new List<BaseActionType>(base.GetActions(userContext, types) ?? new BaseActionType[0]);
            actions.Add(new EditObjectAction(typeof(CRMEntity), "Edit", "", "", () => this, (usercontext, obj) => { ((CRMEntity)obj).Store(); }));
            actions.Add(
                        new YesNoAction("Regenerate Code", "Regenerate Code",
                                        new InvokeServiceMethod(RegenerateDataStructureOfSelectedEntity),
                                        "",
                                        string.Format("Are you sure you want to regenerate code of {0} entity?", crmEntityName))
                        );
            return actions.ToArray();
        }

        private Type GetTypeFromAttributeType(AttributeTypeCode? attributeType)
        {
            if (attributeType != null)
            {
                switch (attributeType)
                {
                    case AttributeTypeCode.BigInt:
                    case AttributeTypeCode.Double:
                        return typeof(double);
                    case AttributeTypeCode.Boolean:
                        return typeof(bool);
                    case AttributeTypeCode.DateTime:
                        return typeof(DateTime);
                    case AttributeTypeCode.Decimal:
                    case AttributeTypeCode.Money:
                        return typeof(decimal);
                    case AttributeTypeCode.Integer:
                        return typeof(int);
                    case AttributeTypeCode.Memo:
                    case AttributeTypeCode.String:
                    case AttributeTypeCode.Picklist:
                        return typeof(string);
                    case AttributeTypeCode.Uniqueidentifier:
                        return typeof(Guid);
                    case AttributeTypeCode.Customer:
                    case AttributeTypeCode.Lookup:
                        return typeof(CRMLookUpTypeField);
                }
            }
            return null;
        }
        public ValidationIssue[] GetValidationIssues()
        {
            List<ValidationIssue> issues = new List<ValidationIssue>();

            if (Connection == null)
            {
                issues.Add(new ValidationIssue(this, "Invalid connection name"));
            }
            else
            {
                string[] logicalNames = Connection.AllEntityNames;
                string[] displayNames = Connection.AllEntityDisplayNames;

                if (logicalNames == null || displayNames == null || logicalNames.Length != displayNames.Length)
                {
                    issues.Add(new ValidationIssue(this, "Error retrieving entity names"));
                }
                else if (!AllEntityNames.Contains(EntityName))
                {
                    issues.Add(new ValidationIssue(this, "Must select a valid entity from the list"));
                }

            }

            return issues.ToArray();
        }
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }


    public class CRMEntityField
    {
        public string FieldName { get; set; }
        public string DisplayName { get; set; }
        public DecisionsNativeType FieldType { get; set; }
        public bool? IsValidForUpdate { get; set; }
        public bool IsRequired { get; set; }
        public string AttributeType { get; set; }
        public CRMOptionsSet[] CRMOptionSet { get; set; }
    }
    public class CRMOptionsSet
    {
        public string OptionName { get; set; }
        public int? OptionValue { get; set; }
    }

    [ValidationRules]
    public class CRMLookUpTypeField
    {
        [PropertyClassification("Look Up Entity Name", 0)]
        [RequiredProperty("Look Up Entity Name is required.")]
        public string LookUpEntityName { get; set; }

        [PropertyClassification("Id", 1)]
        [RequiredProperty("Id is required.")]
        public string Id { get; set; }

    }
}
