using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;

namespace FluentCamlQueries
{
    public static class FluentCamlQueries
    {
        /// <summary>
        /// An example of use:
        /// 
        /// var results = list.Query()
        ///     .When("Fund").IsEqualTo("Fellowships")
        ///     .And.When("Title").Contains("Environmental")
        /// .WithViewContaining("Title")
        ///     .And("Contract").Only
        /// .OrderByAscendingInternalName("Modified", true) // using internal name here
        /// .Finally
        ///     .GetResultsAsList();
        /// </summary>
        public static QueryDefinition Query(this SPList list)
        {
            return new QueryDefinition(list, null);
        }

        public static QueryDefinition QueryInFolder(this SPList list, string subfolderName)
        {
            return new QueryDefinition(list, subfolderName);
        }
        public class QueryDefinition
        {
            public SPList List { get; private set; }
            public List<FieldCondition> Conditions { get; private set; }
            public List<string> Containers { get; private set; }
            public ViewDefinition ViewFields { get; private set; }
            public SPFolder Folder { get; private set; }

            public enum OrderDirection
            {
                Ascending, Descending
            }

            public FieldName OrderByField { get; private set; }
            public OrderDirection OrderByDirection { get; private set; }

            public QueryDefinition(SPList list, string subfolderName)
            {
                List = list;
                Conditions = new List<FieldCondition>();
                Containers = new List<string>();
                ViewFields = new ViewDefinition(this);
                if (string.IsNullOrEmpty(subfolderName) == false)
                    Folder = list.ParentWeb.GetFolder(string.Format("{0}\\{1}", list.Title, subfolderName));
            }

            public FieldCondition When(string fieldDisplayName)
            {
                var condition = new FieldCondition(this, fieldDisplayName, FieldName.NameType.Display);
                Conditions.Add(condition);
                return condition;
            }

            public FieldCondition WhenInternalName(string fieldName)
            {
                var condition = new FieldCondition(this, fieldName, FieldName.NameType.Internal);
                Conditions.Add(condition);
                return condition;
            }

            public QueryDefinition And
            {
                get
                {
                    Containers.Add("And");
                    return this;
                }
            }

            public QueryDefinition Or
            {
                get
                {
                    Containers.Add("Or");
                    return this;
                }
            }

            public ViewDefinition WithViewContaining(string fieldDisplayName)
            {
                ViewFields.And(fieldDisplayName);
                return ViewFields;
            }

            public ViewDefinition WithViewContainingInternalName(string fieldName)
            {
                ViewFields.AndInternalName(fieldName);
                return ViewFields;
            }

            public QueryDefinition OrderByAscending(string fieldDisplayName)
            {
                OrderByField = new FieldName(fieldDisplayName, FieldName.NameType.Display);
                OrderByDirection = OrderDirection.Ascending;
                return this;
            }

            public QueryDefinition OrderByAscendingInternalName(string fieldName)
            {
                OrderByField = new FieldName(fieldName, FieldName.NameType.Internal);
                OrderByDirection = OrderDirection.Ascending;
                return this;
            }

            public QueryDefinition OrderByDescending(string fieldDisplayName)
            {
                OrderByField = new FieldName(fieldDisplayName, FieldName.NameType.Display);
                OrderByDirection = OrderDirection.Descending;
                return this;
            }

            public QueryDefinition OrderByDescendingInternalName(string fieldName)
            {
                OrderByField = new FieldName(fieldName, FieldName.NameType.Internal);
                OrderByDirection = OrderDirection.Descending;
                return this;
            }

            public QueryFactory Finally
            {
                get { return new QueryFactory(this); }
            }
        }

        public class FieldCondition
        {
            public enum FieldConditionType
            {
                Equals, NotEqual, GreaterThan, GreaterThanOrEqualTo, LessThan, LessThanOrEqualTo, IsNull, IsNotNull, BeginsWith, Contains
            }

            private readonly QueryDefinition queryDefinition;

            public FieldName Field { get; private set; }

            public FieldConditionType ConditionType { get; private set; }
            public object ConditionValue { get; private set; }

            public FieldCondition(QueryDefinition queryDefinition, string fieldName, FieldName.NameType nameType)
            {
                this.queryDefinition = queryDefinition;
                Field = new FieldName(fieldName, nameType);
            }
            
            public QueryDefinition IsEqualTo(object value)
            {
                ConditionType = FieldConditionType.Equals;
                ConditionValue = value;
                return queryDefinition;
            }

            public QueryDefinition IsNotEqualTo(object value)
            {
                ConditionType = FieldConditionType.NotEqual;
                ConditionValue = value;
                return queryDefinition;
            }

            public QueryDefinition IsGreaterThan(object value)
            {
                ConditionType = FieldConditionType.GreaterThan;
                ConditionValue = value;
                return queryDefinition;
            }

            public QueryDefinition IsGreaterThanOrEqualTo(object value)
            {
                ConditionType = FieldConditionType.GreaterThanOrEqualTo;
                ConditionValue = value;
                return queryDefinition;
            }

            public QueryDefinition IsLessThan(object value)
            {
                ConditionType = FieldConditionType.LessThan;
                ConditionValue = value;
                return queryDefinition;
            }

            public QueryDefinition IsLessThanOrEqualTo(object value)
            {
                ConditionType = FieldConditionType.LessThanOrEqualTo;
                ConditionValue = value;
                return queryDefinition;
            }

            public QueryDefinition IsNull()
            {
                ConditionType = FieldConditionType.IsNull;
                return queryDefinition;
            }

            public QueryDefinition IsNotNull()
            {
                ConditionType = FieldConditionType.IsNotNull;
                return queryDefinition;
            }

            public QueryDefinition BeginsWith(object value)
            {
                ConditionType = FieldConditionType.BeginsWith;
                ConditionValue = value;
                return queryDefinition;
            }

            public QueryDefinition Contains(object value)
            {
                ConditionType = FieldConditionType.Contains;
                ConditionValue = value;
                return queryDefinition;
            }
        }

        public class ViewDefinition
        {
            public List<FieldName> FieldNames { get; private set; }

            private readonly QueryDefinition queryDefinition;

            public ViewDefinition(QueryDefinition queryDefinition)
            {
                this.queryDefinition = queryDefinition;
                FieldNames = new List<FieldName>();
            }

            public ViewDefinition And(string fieldName)
            {
                FieldNames.Add(new FieldName(fieldName, FieldName.NameType.Display));
                return this;
            }

            public ViewDefinition AndInternalName(string fieldName)
            {
                FieldNames.Add(new FieldName(fieldName, FieldName.NameType.Internal));
                return this;
            }

            public QueryDefinition Only
            {
                get { return queryDefinition; }
            }
        }

        public class FieldName
        {
            public enum NameType
            {
                Internal, Display
            }

            public string Name { get; private set; }
            public NameType Type { get; private set; }

            public FieldName(string name, NameType type)
            {
                Name = name;
                Type = type;
            }
        }

        public class QueryFactory
        {
            const string baseFieldRef = "<FieldRef Name=\"{0}\" />";
            const string baseOrderBy = "<OrderBy><FieldRef Name=\"{0}\" Ascending=\"{1}\" /></OrderBy>";
            const string baseFieldValue = "<Value Type=\"{0}\">{1}</Value>";

            private readonly QueryDefinition queryDefinition;
            private readonly FieldDetails[] allFieldDetails;

            public QueryFactory(QueryDefinition queryDefinition)
            {
                this.queryDefinition = queryDefinition;
                allFieldDetails =
                    queryDefinition.List.Fields.Cast<SPField>().Select(FieldDetails.FromField).ToArray();
            }

            public QueryCaml GetCaml()
            {
                return new QueryCaml
                    {
                        ViewCaml = ViewCaml(),
                        FilterCaml = FilterCaml(),
                        OrderByCaml = OrderByCaml()
                    };
            }

            public List<SPListItem> GetResultsAsList()
            {
                return ResultsFromQuery(null, false);
            }

            public List<SPListItem> GetResultsAsList(uint? rowLimit)
            {
                return ResultsFromQuery(rowLimit, false);
            }

            public List<SPListItem> GetResultsAsList(uint? rowLimit, bool ignoreFolders)
            {
                return ResultsFromQuery(rowLimit, ignoreFolders);
            }

            private string ViewCaml()
            {
                var viewFieldBuilder = new StringBuilder();
                foreach (var field in queryDefinition.ViewFields.FieldNames)
                    viewFieldBuilder.AppendFormat(baseFieldRef, field.Type == FieldName.NameType.Internal
                        ? field.Name : allFieldDetails.Single(f => f.Title.Equals(field.Name)).InternalName);
                return viewFieldBuilder.ToString();
            }

            private string FilterCaml()
            {
                if (queryDefinition.Conditions.Count == 0)
                    return "<Where></Where>";

                var whereBody = CamlForCondition(queryDefinition.Conditions[0], allFieldDetails);
                for (var i = 0; i < queryDefinition.Containers.Count; i++)
                    whereBody = string.Format("<{0}>{1}{2}</{0}>", 
                        queryDefinition.Containers[i], 
                        whereBody, 
                        CamlForCondition(queryDefinition.Conditions[i + 1], allFieldDetails));

                return "<Where>" + whereBody + "</Where>";
            }

            private string OrderByCaml()
            {
                if (queryDefinition.OrderByField == null)
                    return string.Empty;

                var fieldInternalName = queryDefinition.OrderByField.Type == FieldName.NameType.Internal
                    ? queryDefinition.OrderByField.Name
                    : allFieldDetails.Single(f => f.Title.Equals(queryDefinition.OrderByField.Name)).InternalName;

                return string.Format(baseOrderBy, fieldInternalName, 
                    queryDefinition.OrderByDirection == QueryDefinition.OrderDirection.Ascending ? "TRUE" : "FALSE");
            }

            private List<SPListItem> ResultsFromQuery(uint? rowLimit, bool ignoreFolders)
            {
                return queryDefinition.List.GetItems(GetQuery(rowLimit, ignoreFolders)).Cast<SPListItem>().ToList();
            }

            public SPQuery GetQuery(uint? rowLimit, bool ignoreFolders)
            {
                var query = new SPQuery
                {
                    Query = FilterCaml() + OrderByCaml(),
                    ViewFields = ViewCaml()
                };
                if (rowLimit.HasValue)
                    query.RowLimit = rowLimit.Value;
                if (ignoreFolders)
                    query.ViewAttributes = "Scope='Recursive'";

                if (queryDefinition.Folder != null)
                    query.Folder = queryDefinition.Folder;

                return query;
            }

            private static string CamlForCondition(FieldCondition fieldCondition, IEnumerable<FieldDetails> allFields)
            {
                var condition = ConditionFrom(fieldCondition.ConditionType);

                string fieldInternalName;
                if(fieldCondition.Field.Type == FieldName.NameType.Internal)
                    fieldInternalName = fieldCondition.Field.Name;
                else
                {
                    try
                    {
                        fieldInternalName = allFields.Single(f => f.Title.Equals(fieldCondition.Field.Name)).InternalName;
                    }
                    catch (InvalidOperationException)
                    {
                        throw new DisplayNameNotUniqueException(fieldCondition.Field.Name);
                    }
                }
                
                var fieldRef = string.Format(baseFieldRef, fieldInternalName);

                var fieldValue = string.Empty;
                if (!condition.Equals("IsNull") && !condition.Equals("IsNotNull"))
                {
                    var type = allFields.Single(f => f.InternalName.Equals(fieldInternalName)).Type;
                    fieldValue = string.Format(baseFieldValue,
                        type, Process(fieldCondition.ConditionValue, type));
                }

                return string.Format("<{0}>{1}{2}</{0}>", condition, fieldRef, fieldValue);
            }

            private static string ConditionFrom(FieldCondition.FieldConditionType fieldConditionType)
            {
                switch (fieldConditionType)
                {
                    case FieldCondition.FieldConditionType.NotEqual:
                        return "Neq";
                    case FieldCondition.FieldConditionType.GreaterThan:
                        return "Gt";
                    case FieldCondition.FieldConditionType.GreaterThanOrEqualTo:
                        return "Geq";
                    case FieldCondition.FieldConditionType.LessThan:
                        return "Lt";
                    case FieldCondition.FieldConditionType.LessThanOrEqualTo:
                        return "Leq";
                    case FieldCondition.FieldConditionType.IsNull:
                        return "IsNull";
                    case FieldCondition.FieldConditionType.IsNotNull:
                        return "IsNotNull";
                    case FieldCondition.FieldConditionType.BeginsWith:
                        return "BeginsWith";
                    case FieldCondition.FieldConditionType.Contains:
                        return "Contains";
                    default:
                        return "Eq";
                }
            }

            private static string Process(object conditionValue, IEquatable<string> type)
            {
                if(type.Equals("DateTime"))
                {
                    var value = conditionValue.GetType().Equals(typeof (DateTime))
                                    ? (DateTime) conditionValue
                                    : DateTime.Parse(conditionValue.ToString());
                    return SPUtility.CreateISO8601DateTimeFromSystemDateTime(value);
                }

                if (type.Equals("Boolean"))
                    return (bool) conditionValue ? "TRUE" : "FALSE";

                return conditionValue.ToString();
            }

            class FieldDetails
            {
                public string Title { get; private set; }
                public string InternalName { get; private set; }
                public string Type { get; private set; }

                public static FieldDetails FromField(SPField field)
                {
                    return new FieldDetails
                        {
                            Title = field.Title,
                            InternalName = field.InternalName,
                            Type = field.Type.ToString()
                        };
                }
            }
        }

        public class QueryCaml
        {
            public string ViewCaml { get; set; }
            public string FilterCaml { get; set; }
            public string OrderByCaml { get; set; }
        }

        public class DisplayNameNotUniqueException : Exception
        {
            public DisplayNameNotUniqueException(string fieldName) : 
                base(string.Format("There is more than one field with the display name {0}. Try using this fields internal name instead.", fieldName))
            { }
        }
    }
}