using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics.CodeAnalysis;
using System.Management.Automation;
using System.Security;
using DbExtensions;
using SPEventCache;

namespace GetSPEventCache
{
    [Cmdlet(VerbsCommon.Get, "SPEventCache", SupportsShouldProcess = true)]
    [SuppressMessage("ReSharper", "UseStringInterpolation")]
    public class GetSPEventCache : PSCmdlet
    {
        private string _databaseName;
        private string _databaseServer;
        private string _userName;
        private SecureString _passWord;
        private DateTime _eventStart = DateTime.MinValue;
        private DateTime _eventEnd = DateTime.MaxValue;
        private string _modifiedBy = string.Empty;
        private List<Events> _data = new List<Events>();
        private Sort _sortOrder = Sort.NoOrder;
        private int _topRecords = 0;
        private Fields _orderBy = Fields.NoField;
        private string _correlationId = string.Empty;
        private string _webId = string.Empty;
        private string _siteId = string.Empty;
        private string _listId = string.Empty;
        private string _itemFullUrl = string.Empty;
        private string _itemName = string.Empty;
        private string _docId = string.Empty;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (_passWord != null)
            {
                
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            var connectionString = new SqlConnectionStringBuilder
            {
                DataSource = _databaseServer,
                InitialCatalog = _databaseName
            };

            if (_userName != null && _passWord != null)
            {
                connectionString.UserID = _userName;

                var password = DecryptSecurePassword.SecurePasswordToString(_passWord);

                connectionString.Password = password;
            }
            else
            {
                connectionString.IntegratedSecurity = true;
            }

            var query = DynamicSqlBuilder(_eventStart, _eventEnd, _correlationId, _orderBy, _sortOrder, 
                _siteId, _webId, _listId, _docId, _itemName, _itemFullUrl, _modifiedBy, _topRecords);

            using (var connection = new SqlConnection(connectionString.ToString()))
            {
                connection.Open();
                using (var command = new SqlCommand(query.ToString(), connection))
                {
                    var table = new DataTable();
                    var adapter = new SqlDataAdapter(command);
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        var eventTime = (row.ItemArray[0] as DateTime?) ?? DateTime.MinValue;
                        var iD = (Convert.IsDBNull(row.ItemArray[1])) ? 0 : Convert.ToInt32(row.ItemArray[1]);
                        var siteId = (row.ItemArray[2] as Guid?) ?? Guid.Empty;
                        var webId = (row.ItemArray[3] as Guid?) ?? Guid.Empty;
                        var listId = (row.ItemArray[4] as Guid?) ?? Guid.Empty;
                        var itemId = (Convert.IsDBNull(row.ItemArray[5])) ? 0 : Convert.ToInt32(row.ItemArray[5]);
                        var docId = (row.ItemArray[6] as Guid?) ?? Guid.Empty;
                        var guid0 = (row.ItemArray[7] as Guid?) ?? Guid.Empty;
                        var int0 = (Convert.IsDBNull(row.ItemArray[8])) ? 0 : Convert.ToInt32(row.ItemArray[8]);
                        var int1 = (Convert.IsDBNull(row.ItemArray[9])) ? 0 : Convert.ToInt32(row.ItemArray[9]);
                        var contentTypeId = row.ItemArray[10] == DBNull.Value ? null : (byte[])row.ItemArray[10];
                        var itemName = (row.ItemArray[11] == null) ? string.Empty : row.ItemArray[11].ToString();
                        var itemFullUrl = (row.ItemArray[12] == null) ? string.Empty : row.ItemArray[12].ToString();
                        var eventType = (Convert.IsDBNull(row.ItemArray[13])) ? 0 : Convert.ToInt32(row.ItemArray[13]);
                        var objectType = (Convert.IsDBNull(row.ItemArray[14])) ? 0 : Convert.ToInt32(row.ItemArray[14]);
                        var modifiedBy = (row.ItemArray[15] == null) ? string.Empty : row.ItemArray[15].ToString();
                        var timeLastModified = (row.ItemArray[16] as DateTime?) ?? DateTime.MinValue;
                        var eventData = row.ItemArray[17] == DBNull.Value ? null : (byte[])row.ItemArray[17];
                        var acl = row.ItemArray[18] == DBNull.Value ? null : (byte[])row.ItemArray[18];
                        var docClientId = row.ItemArray[19] == DBNull.Value ? null : (byte[])row.ItemArray[19];
                        var correlationId = (row.ItemArray[20] as Guid?) ?? Guid.Empty;

                        _data.Add(new Events(eventTime, iD, siteId, webId, listId, itemId, docId, guid0, int0, int1, contentTypeId,
                            itemName, itemFullUrl, eventType, objectType, modifiedBy, timeLastModified, eventData, acl, docClientId,
                            correlationId));
                    }
                }
                connection.Close();
                WriteObject(_data, true);
            }
        }

        SqlBuilder DynamicSqlBuilder(DateTime? eventStart, DateTime? eventEnd, string correlationId, Fields? orderBy, Sort? sortOrder, 
            string siteId, string webId, string listId, string docId, string itemName, string itemFullUrl, string modifiedBy, int topRecords)
        {
            var correlationIdGuid = ValidateGuid(correlationId);
            var siteIdGuid = ValidateGuid(siteId);
            var webIdGuid = ValidateGuid(webId);
            var listIdGuid = ValidateGuid(listId);
            var docIdGuid = ValidateGuid(docId);

            var topQuery = topRecords > 0 ? string.Format("TOP({0}) * ", _topRecords) : @"* ";
            var orderQuery = string.Empty;

            if (orderBy.HasValue && orderBy != Fields.NoField)
            {
                orderQuery = string.Format("{0}", orderBy.Value);

                if (sortOrder.HasValue && sortOrder != Sort.NoOrder)
                {
                    orderQuery = string.Format("{0} {1}", orderQuery, sortOrder.Value);
                }
            }

            var query = SQL
                .SELECT(topQuery)
                .FROM("EventCache ")
                .WHERE()
                ._If(eventStart.HasValue && eventStart != DateTime.MinValue,
                    string.Format("EventTime > '{0}'", eventStart))
                ._If(eventEnd.HasValue && eventEnd != DateTime.MaxValue, string.Format("EventTime < '{0}'", eventEnd))
                ._If(correlationIdGuid != Guid.Empty, string.Format("CorrelationId = '{0}'", correlationIdGuid))
                ._If(siteIdGuid != Guid.Empty, string.Format("SiteId = '{0}'", siteIdGuid))
                ._If(webIdGuid != Guid.Empty, string.Format("WebId = '{0}'", webIdGuid))
                ._If(listIdGuid != Guid.Empty, string.Format("ListId = '{0}'", listIdGuid))
                ._If(docIdGuid != Guid.Empty, string.Format("DocId = '{0}'", docIdGuid))
                ._If(!string.IsNullOrEmpty(itemName), string.Format("ItemName LIKE '%{0}%'", itemName))
                ._If(!string.IsNullOrEmpty(itemFullUrl), string.Format("ItemFullUrl LIKE '%{0}%'", itemFullUrl))
                ._If(!string.IsNullOrEmpty(modifiedBy), string.Format("ModifiedBy LIKE '%{0}%'", modifiedBy));

            if (!string.IsNullOrEmpty(orderQuery))
            {
                query.ORDER_BY(orderQuery);
            }

            return query;
        }

        static Guid ValidateGuid(string guidValue)
        {
            if (string.IsNullOrEmpty(guidValue)) return Guid.Empty;

            try
            {
                Guid guidResult;
                Guid.TryParse(guidValue, out guidResult);
                return guidResult;
            }
            catch
            {
                return Guid.Empty;
            }
        }


        #region EventCacheEnums

        [Flags]
        public enum EventObjectTypes : uint
        {
            NoObjectType = 0,
            ListItem = 0x00000001, 
            List = 0x00000002, 
            Site = 0x00000004, 
            SiteCollection = 0x00000008, 
            File = 0x00000010, 
            Folder = 0x00000020, 
            Alert = 0x00000040, 
            User = 0x00000080, 
            Group = 0x00000100, 
            ContentType = 0x00000200, 
            Field = 0x00000400, 
            SecurityPolicy = 0x00000800, 
            View = 0x000001000 
        }

        [Flags]
        public enum EventTypes : uint
        {
            NoEventType = 0,
            ListItem = 0x00000001, 
            ListItemModified = 0x00000002, 
            ListItemDeleted = 0x00000004, 
            ListItemRestored = 0x00000008, 
            DiscussionListItemAdded = 0x00000010, 
            DiscussionListItemModified = 0x00000020, 
            DiscussionListItemDeleted = 0x00000040, 
            DiscussionListItemClosed = 0x00000080, 
            DiscussionListItemActivated = 0x00000100, 
            GenericAddEvent = 0x000001000, 
            GenericModificationEvent = 0x000002000, 
            GenericDeleteEvent = 0x000004000, 
            GenericRenameEvent = 0x000008000, 
            MoveInto = 0x000010000, 
            Restore = 0x000020000, 
            PermissionLevelAdded = 0x000040000, 
            RoleAssignmentAdded = 0x000080000, 
            ModificationExtendedBySystem = 0x000100000, 
            MemberAddedToGroup = 0x000200000, 
            MemberDeletedFromGroup = 0x000400000, 
            PermissionLevelDeleted = 0x000800000, 
            PermissionLevelUpdated = 0x001000000, 
            RoleAssignmentDeleted = 0x002000000, 
            MoveAway = 0x004000000, 
            NavigationStructureChanged = 0x008000000
        }

        #endregion

        #region Enums
        public enum Sort
        {
            Desc,
            Asc,
            NoOrder
        };
        
        public enum Fields
        {
            EventTime,
            Id,
            SiteId,
            WebId,
            ListId,
            ItemId,
            DocId,
            ItemFullUrl,
            EventType,
            ObjectType,
            ModifiedBy,
            TimeLastModified,
            DocClientId,
            NoField
        }
        #endregion
        #region Properties

        [Parameter(Mandatory = true, Position = 0, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public string DatabaseName
        {
            get { return _databaseName;}
            set { _databaseName = value; }
        }

        [Parameter(Mandatory = true, Position = 1, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public string DatabaseServer
        {
            get { return _databaseServer;}
            set { _databaseServer = value; }
        }

        [Parameter(Mandatory = false, Position = 2)]
        [ValidateNotNullOrEmpty]
        public DateTime EventStart
        {
            get { return _eventStart;}
            set { _eventStart = value; }
        }

        [Parameter(Mandatory = false, Position = 3)]
        [ValidateNotNullOrEmpty]
        public DateTime EventEnd
        {
            get { return _eventEnd; }
            set { _eventEnd = value; }
        }
        
        [Parameter(Mandatory = false, Position = 4)]
        [ValidateNotNullOrEmpty]
        public Fields OrderBy
        {
            get { return _orderBy; }
            set { _orderBy = value; }
        }
        
        [Parameter(Mandatory = false, Position = 5)]
        [ValidateNotNullOrEmpty]
        public Sort SortOrder
        {
            get { return _sortOrder; }
            set { _sortOrder = value; }
        }

        [Parameter(Mandatory = false, Position = 6)]
        [ValidateNotNullOrEmpty]
        public int TopRecords
        {
            get { return _topRecords; }
            set { _topRecords = value; }
        }

        [Parameter(Mandatory = false, Position = 7)]
        [ValidateNotNullOrEmpty]
        public string CorrelationId
        {
            get { return _correlationId; }
            set { _correlationId = value; }
        }

        [Parameter(Mandatory = false, Position = 8)]
        [ValidateNotNullOrEmpty]
        public string WebId
        {
            get { return _webId;}
            set { _webId = value; }
        }

        [Parameter(Mandatory = false, Position = 9)]
        [ValidateNotNullOrEmpty]
        public string SiteId
        {
            get { return _siteId; }
            set { _siteId = value; }
        }

        [Parameter(Mandatory = false, Position = 10)]
        [ValidateNotNullOrEmpty]
        public string ListId
        {
            get { return _listId; }
            set { _listId = value; }
        }

        [Parameter(Mandatory = false, Position = 11)]
        [ValidateNotNullOrEmpty]
        public string DocId
        {
            get { return _docId;}
            set { _docId = value; }
        }

        [Parameter(Mandatory = false, Position = 12)]
        [ValidateNotNullOrEmpty]
        public string ItemName
        {
            get { return _itemName; }
            set { _itemName = value; }
        }

        [Parameter(Mandatory = false, Position = 13)]
        [ValidateNotNullOrEmpty]
        public string ItemFullUrl
        {
            get { return _itemFullUrl; }
            set { _itemFullUrl = value; }
        }

        [Parameter(Mandatory = false, Position = 14)]
        [ValidateNotNullOrEmpty]
        public string ModifiedBy
        {
            get { return _modifiedBy; }
            set { _modifiedBy = value; }
        }

        [Parameter(Mandatory = false)]
        [ValidateNotNullOrEmpty]
        public string Username
        {
            get { return _userName; }
            set { _userName = value; }
        }

        [Parameter(Mandatory = false)]
        [ValidateNotNullOrEmpty]
        public SecureString Password
        {
            get { return _passWord; }
            set { _passWord = value; }
        }

        #endregion
    }

    public class Events
    {
        public DateTime EventTime { get; set; }
        public int ID { get; set; }
        public Guid SiteId { get; set; }
        public Guid WebId { get; set; }
        public Guid ListId { get; set; }
        public int ItemId { get; set; }
        public Guid DocId { get; set; }
        public Guid? Guid0 { get; set; }
        public int Int0 { get; set; }
        public int Int1 { get; set; }
        public byte[] ContentTypeId { get; set; }
        public string ItemName { get; set; }
        public string ItemFulLUrl { get; set; }
        public object EventType { get; set; }
        public object ObjectType { get; set; }
        public string ModifiedBy { get; set; }
        public DateTime TimeLastModified { get; set; }
        public byte[] EventData { get; set; }
        public byte[] ACL { get; set; }
        public byte[] DocClientId { get; set; }
        public Guid CorrelationId { get; set; }

        public Events() { }
        public Events(DateTime eventTime, int iD, Guid siteId, Guid webId, Guid listId, int itemId, 
            Guid docId, Guid? guid0, int int0, int int1, byte[] contentTypeId, string itemName,
            string itemFullUrl, int eventType, int objectType, string modifiedBy, DateTime timeLastModified,
            byte[] eventData, byte[] acl, byte[] docClientId, Guid correlationId)
        {
            EventTime = eventTime;
            ID = iD;
            SiteId = siteId;
            WebId = webId;
            ListId = listId;
            ItemId = itemId;
            DocId = docId;
            Guid0 = guid0;
            Int0 = int0;
            Int1 = int1;
            ContentTypeId = contentTypeId;
            ItemName = itemName;
            ItemFulLUrl = itemFullUrl;

            var evtTypes = Enum.Parse(typeof(GetSPEventCache.EventTypes), Convert.ToString(eventType));
            EventType = evtTypes;

            var objTypes = Enum.Parse(typeof (GetSPEventCache.EventObjectTypes), Convert.ToString(objectType));
            ObjectType = objTypes;

            ModifiedBy = modifiedBy;
            TimeLastModified = timeLastModified;
            EventData = eventData;
            ACL = acl;
            DocClientId = docClientId;
            CorrelationId = correlationId;
        }
    }
}
