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
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Decisions.MSCRM
{
    [Writable]
    [ORMEntity]
    [ValidationRules]
    public class CRMConnection : BaseORMEntity, INotifyPropertyChanged, IValidationSource
    {
        #region Fields
        [WritableValue]
        [ORMPrimaryKeyField]
        [PropertyHidden]
        public string connectionId;


        [WritableValue]
        [ORMField] // todo: this needs to be unique.
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

        [WritableValue]
        [ORMField(typeof(ORMXmlSerializedFieldConverter))]
        private string[] allEntityNames;

        [WritableValue]
        [ORMField(typeof(ORMXmlSerializedFieldConverter))]
        private string[] allEntityDisplayNames;

        #endregion
        #region Properties

        [PropertyClassification("Connection Name", 10)]
        [RequiredProperty("Connection Name is required.")]
        public string ConnectionName
        {
            get
            {
                return connectionName;
            }
            set
            {
                connectionName = value;
                OnPropertyChanged();
            }
        }

        [PropertyClassification("Override Connection String", 15)]
        public bool OverrideConnectionString
        {
            get { return overrideConnectionString; }
            set { overrideConnectionString = value; OnPropertyChanged(); }
        }

        [PropertyClassification("Connection String", 16)]
        [BooleanPropertyHidden(nameof(OverrideConnectionString), false)]
        public string ConnectionString
        {
            get { return connectionString; }
            set { connectionString = value; OnPropertyChanged(); }
        }

        [PropertyClassification("Organization Url", 20)]
        [BooleanPropertyHidden(nameof(OverrideConnectionString), true)]
        public string OrganisationUrl
        {
            get
            {
                return organisationUrl;
            }
            set
            {
                organisationUrl = value;
                OnPropertyChanged();
            }
        }

        [PropertyClassification("Domain", 30)]
        [BooleanPropertyHidden(nameof(OverrideConnectionString), true)]
        public string Domain
        {
            get
            {
                return domain;
            }
            set
            {
                domain = value;
                OnPropertyChanged();
            }
        }

        [PropertyClassification("Username", 40)]
        [BooleanPropertyHidden(nameof(OverrideConnectionString), true)]
        public string UserName
        {
            get
            {
                return userName;
            }
            set
            {
                userName = value;
                OnPropertyChanged();
            }
        }
        [PropertyClassification("Password", 50)]
        [PasswordText]
        [ExcludeInDescription]
        [BooleanPropertyHidden(nameof(OverrideConnectionString), true)]
        public string Password
        {
            get
            {
                return password;
            }
            set
            {
                password = value;
                OnPropertyChanged();
            }
        }

        [PropertyHidden]
        public string[] AllEntityNames
        {
            get
            {
                return allEntityNames;
            }
            set
            {
                allEntityNames = value;
            }
        }

        [PropertyHidden]
        public string[] AllEntityDisplayNames
        {
            get
            {
                return allEntityDisplayNames;
            }
            set
            {
                allEntityDisplayNames = value;
            }
        }

        #endregion
        internal string GetConnectionString()
        {
            if (overrideConnectionString)
                return connectionString;
            return string.Format("Url={0}; domain={1}; Username={2}; Password={3}; RequireNewInstance=True;", OrganisationUrl, Domain, UserName, Password);
        }

        internal static CRMConnection GetCRMConnectionById(string connectionId)
        {
            CRMConnection connection = ModuleSettingsAccessor<CRMSettings>.Instance.Connections.FirstOrDefault(x => x.connectionId == connectionId);
            return connection;
        }

        internal static CRMConnection GetCRMConnectionForName(string connectionName)
        {
            CRMConnection connection = ModuleSettingsAccessor<CRMSettings>.Instance.Connections.FirstOrDefault(x => x.connectionName == connectionName);
            return connection;
        }

        public override string ToString()
        {
            if (overrideConnectionString)
                return $"CRM Connection ({connectionString})";
            string cName = ConnectionName ?? "(no name)";
            string uName = UserName ?? "(no user)";
            string url = OrganisationUrl ?? "(no url)";
            return string.Format("{0}  ({1} at {2})", cName, uName, url);
        }

        public override void BeforeSave()
        {
            CRMConnection otherConn = ModuleSettingsAccessor<CRMSettings>.Instance.Connections
                .FirstOrDefault(x => x.connectionId != this.connectionId && x.connectionName == this.connectionName);
            if (otherConn != null) throw new InvalidOperationException("Another connection already exists with this name, please choose another name.");
            base.BeforeSave();
            RetrieveEntityList();
        }

        public void RetrieveEntityList()
        {
            try
            {
                Log log = new Log("CRMConnection");
                try
                {
                    string trimmedConnectionString = GetConnectionString();
                    if (trimmedConnectionString.Contains("Password"))
                    {
                        trimmedConnectionString = trimmedConnectionString.Substring(0, trimmedConnectionString.IndexOf("Password"));
                    }
                    log.Debug($"Making connection for '{this.connectionName}' ({this.connectionId}) with connection info \"{trimmedConnectionString}\".");
                }
                catch { } // Swallow log exceptions

                CrmServiceClient conn = new CrmServiceClient(GetConnectionString());
                IOrganizationService serviceProxy = conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
                if (serviceProxy != null)
                {
                    RetrieveAllEntitiesRequest req = new RetrieveAllEntitiesRequest()
                    {
                        EntityFilters = EntityFilters.Entity
                    };
                    RetrieveAllEntitiesResponse res = (RetrieveAllEntitiesResponse)serviceProxy.Execute(req);
                    this.allEntityNames = res.EntityMetadata.Select(x => x.LogicalName).ToArray();
                    this.allEntityDisplayNames = res.EntityMetadata.Select(x =>
                    {
                        string displayName;
                        if (x.DisplayName.LocalizedLabels != null && x.DisplayName.LocalizedLabels.Count > 0)
                        {
                            displayName = x.DisplayName.LocalizedLabels[0].Label;
                        }
                        else displayName = x.LogicalName;
                        return displayName;
                    }).ToArray();
                }
            }
            catch(System.IO.FileNotFoundException)
            {
                // FileNotFound is sometimes thrown from within Microsoft.Xrm.Tooling.Connector, probably when the connection info is invalid.
                this.allEntityNames = null;
                this.allEntityDisplayNames = null;
                throw new InvalidOperationException("Connection could not be made - check connection info or retry.");
            }
        }

        internal string GetCrmConnectionFolderId()
        {
            if (connectionId == null) throw new InvalidOperationException("This connection has no ID and cannot be used to create a folder.");
            return CRMEntity.CRM_GENERATED_TYPES_FOLDER_ID + connectionId;
        }

        internal void EnsureCrmConnectionFolderExists()
        {
            if (connectionId == null) throw new InvalidOperationException("Cannot create a folder for an incomplete connection.");
            ORM<Folder> folderORM = new ORM<Folder>();
            string crmConnectionFolderId = GetCrmConnectionFolderId();
            Folder crmEntityFolder = folderORM.Fetch(crmConnectionFolderId);
            if (crmEntityFolder == null)
            {
                crmEntityFolder = new Folder(crmConnectionFolderId, connectionName, CRMEntity.CRM_GENERATED_TYPES_FOLDER_ID);
                folderORM.Store(crmEntityFolder);
            }
        }

        public ValidationIssue[] GetValidationIssues()
        {
            List<ValidationIssue> issues = new List<ValidationIssue>();

            if (OverrideConnectionString)
            {
                if (string.IsNullOrEmpty(ConnectionString))
                    issues.Add(new ValidationIssue(this, "Connection string is required", "", BreakLevel.Fatal, nameof(ConnectionString)));
            }
            else
            {
                if (string.IsNullOrEmpty(OrganisationUrl))
                    issues.Add(new ValidationIssue(this, "Organization URL is required", "", BreakLevel.Fatal, nameof(OrganisationUrl)));
                if (string.IsNullOrEmpty(Domain))
                    issues.Add(new ValidationIssue(this, "Domain is required", "", BreakLevel.Fatal, nameof(Domain)));
                if (string.IsNullOrEmpty(UserName))
                    issues.Add(new ValidationIssue(this, "Username is required", "", BreakLevel.Fatal, nameof(UserName)));
                if (string.IsNullOrEmpty(Password))
                    issues.Add(new ValidationIssue(this, "Password is required", "", BreakLevel.Fatal, nameof(Password)));
            }

            return issues.ToArray();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
