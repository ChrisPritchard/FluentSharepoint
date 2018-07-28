using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Reflection;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Navigation;

namespace FluentSharePoint
{
    public static partial class FluentCreationExtensions
    {
        public static ContentTypeDefinition EnsureContentType(this SPWeb web, string name)
        {
            return new ContentTypeDefinition(web, name);
        }

        public static ListDefinition EnsureList(this SPWeb web, string name)
        {
            return new ListDefinition(web, name);
        }

        public static ListDefinition EnsureHiddenList(this SPWeb web, string name)
        {
            return new ListDefinition(web, name) { IsHidden = true };
        }

        public static RoleDefinition EnsureRoleDefinition(this SPWeb web, string name)
        {
            return new RoleDefinition(web, name);
        }

        public static GroupDefinition EnsureUserGroup(this SPWeb web, string name)
        {
            return new GroupDefinition(web, name);
        }

        public static RoleAssociation AddRole(this SPWeb web, string name)
        {
            return new RoleAssociation(web, web.RoleAssignments, name);
        }

        public static RoleAssociation AddRole(this SPList list, string name)
        {
            return new RoleAssociation(list.ParentWeb, list.RoleAssignments, name);
        }

        public static RoleAssociation AddRole(this SPListItem item, string name)
        {
            return new RoleAssociation(item.ParentList.ParentWeb, item.RoleAssignments, name);
        }

        public partial interface ICanHaveFields
        {
            SPWeb Web { get; }
        }

        public interface ICanHaveRoles
        { }

        public partial class ContentTypeDefinition : ICanHaveFields
        {
            public SPWeb Web { get; private set; }
            public string Name { get; private set; }
            public string ParentContentType { get; protected set; }
            public List<FieldDefinition<ContentTypeDefinition>> FieldDefinitions { get; protected set; }
            public string NewFormUrl { get; protected set; }

            public ContentTypeDefinition(SPWeb web, string name)
            {
                Web = web;
                Name = name;
                ParentContentType = "Item";
                FieldDefinitions = new List<FieldDefinition<ContentTypeDefinition>>();
                NewFormUrl = string.Empty;
            }

            public ContentTypeDefinition FromType(string parentContentTypeName)
            {
                ParentContentType = parentContentTypeName;
                return this;
            }

            public ContentTypeDefinition With(Func<ContentTypeDefinition, ContentTypeDefinition> settingsMethod)
            {
                return settingsMethod(this);
            }

            public ContentTypeDefinition WithNewFormUrl(string url)
            {
                NewFormUrl = url;
                return this;
            }

            public FieldDefinition<ContentTypeDefinition> WithField(string fieldName)
            {
                var definition = new FieldDefinition<ContentTypeDefinition>(fieldName, this);
                FieldDefinitions.Add(definition);
                return definition;
            }

            public FieldDefinition<ContentTypeDefinition> WithHiddenField(string fieldName)
            {
                var definition = new FieldDefinition<ContentTypeDefinition>(fieldName, this) { IsHidden = true, IsRequired = false };
                FieldDefinitions.Add(definition);
                return definition;
            }

            public FieldDefinition<ContentTypeDefinition> WithOptionalField(string fieldName)
            {
                var definition = new FieldDefinition<ContentTypeDefinition>(fieldName, this) { IsRequired = false };
                FieldDefinitions.Add(definition);
                return definition;
            }

            public SPContentType CreateAsPartOfGroup(string groupName)
            {
                if (Web.ContentTypes[Name] != null)
                    return Web.ContentTypes[Name];

                var contentType = new SPContentType(Web.AvailableContentTypes[ParentContentType], Web.ContentTypes, Name) { Group = groupName };

                foreach (var fieldInfo in FieldDefinitions.Select(fieldDef => new { fieldDef.IsHidden, Field = fieldDef.CreateIn(Web.Fields, groupName) }))
                    contentType.FieldLinks.Add(new SPFieldLink(fieldInfo.Field) { Hidden = fieldInfo.IsHidden });

                if (!string.IsNullOrEmpty(NewFormUrl))
                    contentType.NewFormUrl = NewFormUrl;

                Web.ContentTypes.Add(contentType);
                Web.Update();

                contentType.Update();
                return contentType;
            }
        }

        public partial class ListDefinition : ICanHaveFields, ICanHaveRoles
        {
            public SPWeb Web { get; private set; }

            private readonly string name;
            private string description;
            private SPListTemplateType type;

            private List<string> contentTypeNames;
            private readonly List<string> singleFieldIndexes;
            private readonly List<string[]> compoundFieldIndexes;
            public bool IsHidden { private get; set; }
            private string[] defaultView;
            private string quickLinksHeader;

            private readonly List<FieldDefinition<ListDefinition>> fieldDefinitions;
            private readonly Dictionary<SPEventReceiverType, Type> events;
            private readonly List<RoleAssociation<ListDefinition>> roleAssociations;
            private readonly List<FolderDefinition> folderDefinitions;

            public ListDefinition(SPWeb web, string name)
            {
                Web = web;
                this.name = name;
                description = string.Empty;
                type = SPListTemplateType.GenericList;

                contentTypeNames = new List<string>();
                singleFieldIndexes = new List<string>();
                compoundFieldIndexes = new List<string[]>();
                IsHidden = false;

                fieldDefinitions = new List<FieldDefinition<ListDefinition>>();
                events = new Dictionary<SPEventReceiverType, Type>();
                roleAssociations = new List<RoleAssociation<ListDefinition>>();
                folderDefinitions = new List<FolderDefinition>();
            }

            public ListDefinition OfType(SPListTemplateType listType)
            {
                type = listType;
                return this;
            }

            public ListDefinition WithDescription(string listDescription)
            {
                description = listDescription;
                return this;
            }

            public ListDefinition WithContentType(string contentTypeName)
            {
                contentTypeNames.Add(contentTypeName);
                return this;
            }

            public ListDefinition WithContentTypes(params string[] contentTypeNameList)
            {
                contentTypeNames.AddRange(contentTypeNameList);
                return this;
            }

            public ListDefinition WithFieldIndex(string fieldInternalName)
            {
                singleFieldIndexes.Add(fieldInternalName);
                return this;
            }

            public ListDefinition WithFieldIndexes(params string[] fieldNames)
            {
                singleFieldIndexes.AddRange(fieldNames);
                return this;
            }

            public ListDefinition WithCompoundFieldIndex(string firstFieldName, string secondFieldName)
            {
                compoundFieldIndexes.Add(new[] {firstFieldName, secondFieldName});
                return this;
            }

            public ListDefinition WithViewContaining(params string[] displayedFieldInternalNames)
            {
                defaultView = displayedFieldInternalNames;
                return this;
            }

            public ListDefinition AndEvent<TEventClass>(SPEventReceiverType eventType)
            {
                if (!events.ContainsKey(eventType))
                    events.Add(eventType, typeof(TEventClass));
                return this;
            }

            public ListDefinition ShowOnQuickLinksUnder(string parentNodeName)
            {
                quickLinksHeader = parentNodeName;
                return this;
            }

            public FieldDefinition<ListDefinition> WithField(string fieldName)
            {
                var definition = new FieldDefinition<ListDefinition>(fieldName, this);
                fieldDefinitions.Add(definition);
                return definition;
            }

            public FieldDefinition<ListDefinition> WithHiddenField(string fieldName)
            {
                var definition = new FieldDefinition<ListDefinition>(fieldName, this) { IsHidden = true, IsRequired = false };
                fieldDefinitions.Add(definition);
                return definition;
            }

            public FieldDefinition<ListDefinition> WithOptionalField(string fieldName)
            {
                var definition = new FieldDefinition<ListDefinition>(fieldName, this) { IsRequired = false };
                fieldDefinitions.Add(definition);
                return definition;
            }

            public RoleAssociation<ListDefinition> WithRole(string roleName)
            {
                var association = new RoleAssociation<ListDefinition>(this, roleName);
                roleAssociations.Add(association);
                return association;
            }

            public FolderDefinition WithFolder(string folderName)
            {
                var definition = new FolderDefinition(this, folderName);
                folderDefinitions.Add(definition);
                return definition;
            }

            public SPList CreateAndVerify()
            {
                return CreateAndOrVerify(true);
            }

            public SPList Create()
            {
                return CreateAndOrVerify(false);
            }

            private SPList CreateAndOrVerify(bool verify)
            {
                var list = Web.Lists.Cast<SPList>().SingleOrDefault(l => l.Title.Equals(name));
                if (list != null && !verify)
                    return list;
                
                if(list == null)
                {
                    var listId = Web.Lists.Add(name, description, type);
                    list = Web.Lists[listId];
                }

                contentTypeNames = contentTypeNames.Distinct().ToList();

                var existingTypes = list.ContentTypes.Cast<SPContentType>().Select(c => c.Name).ToArray();
                contentTypeNames = contentTypeNames.Where(c => !existingTypes.Contains(c)).ToList();

                list.ContentTypesEnabled = true;
                foreach (var contentTypeName in contentTypeNames)
                    list.ContentTypes.Add(Web.AvailableContentTypes[contentTypeName]);

                if (contentTypeNames.Count > 0 && existingTypes[0].Equals("Item"))
                    list.ContentTypes[0].Delete();

                list.Hidden = IsHidden;
                list.Update();

                if (defaultView != null)
                    SetView(list, defaultView);

                if (!string.IsNullOrEmpty(quickLinksHeader))
                {
                    var parent = Web.Navigation.QuickLaunch.Cast<SPNavigationNode>().FirstOrDefault(n => n.Title.Equals(quickLinksHeader));
                    if (parent != null)
                        parent.Children.AddAsLast(new SPNavigationNode(name, list.RootFolder.ServerRelativeUrl));
                }

                list.EventReceivers.Cast<SPEventReceiverDefinition>().ToList().ForEach(e => e.Delete());
                foreach (var eventDef in events)
                    list.EventReceivers.Add(eventDef.Key, eventDef.Value.Assembly.FullName, eventDef.Value.FullName);

                if (fieldDefinitions.Count > 0)
                {
                    var newFields = fieldDefinitions
                        .Select(definition => new SPFieldLink(definition.CreateIn(list.Fields, null)));

                    foreach (var contentType in list.ContentTypes.Cast<SPContentType>().Where(c => !c.Sealed && !c.Hidden))
                    {
                        foreach (var fieldLink in newFields)
                            contentType.FieldLinks.Add(fieldLink);
                        contentType.Update();
                    }
                    list.Update();
                }

                var nonHiddenSiteFields = Web.Fields.Cast<SPField>().Where(f => !f.Hidden).Select(f => f.InternalName).ToArray();
                var setter = typeof(SPField).GetMethod("SetFieldBoolValue", BindingFlags.NonPublic | BindingFlags.Instance);
                foreach (var field in list.Fields.Cast<SPField>().Where(f => f.Hidden && nonHiddenSiteFields.Contains(f.InternalName)).ToArray())
                {
                    setter.Invoke(field, new object[] { "CanToggleHidden", true });
                    field.Hidden = false;
                    field.Update();
                }

                foreach (var singleFieldIndex in singleFieldIndexes)
                    list.FieldIndexes.Add(list.Fields[singleFieldIndex]);

                foreach (var compoundFieldIndex in compoundFieldIndexes)
                    list.FieldIndexes.Add(list.Fields[compoundFieldIndex[0]], list.Fields[compoundFieldIndex[1]]);

                if (roleAssociations.Count > 0)
                {
                    list.BreakRoleInheritance(false);
                    foreach (var listRoleAssociation in roleAssociations)
                        list.AddRole(listRoleAssociation.RoleDefinitionName).For(listRoleAssociation.PrincipalNames);
                    list.Update();
                }

                foreach (var folderDefinition in folderDefinitions)
                    FolderDefinition.Create(list, folderDefinition);

                return list;
            }

            public SPList CreateWithVersioningAndModeration()
            {
                var list = Create();

                list.EnableModeration = true;
                list.EnableVersioning = true;
                list.Update();

                return list;
            }

            private static void SetView(SPList list, string[] defaultViewFields)
            {
                var defaultView = list.Views.Cast<SPView>().SingleOrDefault(v => v.Title.Equals("Default"));
                if (defaultView != null)
                    list.Views.Delete(defaultView.ID);

                var fields = new StringCollection();
                fields.AddRange(defaultViewFields);
                list.Views.Add("Default", fields, string.Empty, 50, true, true, SPViewCollection.SPViewType.Html, false);
            }

            public class RoleAssociation<TParent> where TParent : ICanHaveRoles
            {
                public string RoleDefinitionName { get; private set; }
                public IEnumerable<string> PrincipalNames { get; private set; }

                private readonly TParent parent;

                public RoleAssociation(TParent parent, string roleDefinitionName)
                {
                    this.parent = parent;
                    RoleDefinitionName = roleDefinitionName;
                }

                public TParent For(params string[] principalNames)
                {
                    PrincipalNames = principalNames;
                    return parent;
                }

                public TParent For(IEnumerable<string> principalNames)
                {
                    PrincipalNames = principalNames;
                    return parent;
                }
            }

            public class FolderDefinition : ICanHaveRoles
            {
                public string Name { get; private set; }
                public ListDefinition Parent { get; private set; }

                private readonly List<RoleAssociation<FolderDefinition>> roleAssociations;

                public FolderDefinition(ListDefinition parent, string name)
                {
                    Name = name;
                    Parent = parent;

                    roleAssociations = new List<RoleAssociation<FolderDefinition>>();
                }

                public RoleAssociation<FolderDefinition> WithRole(string roleName)
                {
                    var association = new RoleAssociation<FolderDefinition>(this, roleName);
                    roleAssociations.Add(association);
                    return association;
                }

                public ListDefinition And { get { return Parent; } }

                public ListDefinition Finally { get { return Parent; } }

                public static void Create(SPList list, FolderDefinition definition)
                {
                    var web = list.ParentWeb;
                    var existingFolder = web.GetFolder(list.Title + "/" + definition.Name);
                    if (existingFolder.Exists)
                        return;

                    var newFolder = list.Items.Add(string.Empty, SPFileSystemObjectType.Folder, definition.Name);
                    newFolder.Update();

                    if (definition.roleAssociations.Count <= 0)
                        return;

                    newFolder.BreakRoleInheritance(false);
                    foreach (var association in definition.roleAssociations)
                        newFolder.AddRole(association.RoleDefinitionName).For(association.PrincipalNames);
                }
            }
        }

        public partial class FieldDefinition<TParentDefinition> where TParentDefinition : ICanHaveFields
        {
            public string Name { get; private set; }
            private readonly TParentDefinition parent;
            private SPFieldType type;
            private bool shouldRenderAsRichText;

            private string displayName;
            public bool IsRequired { private get; set; }
            public bool IsHidden { get; set; }

            private StringCollection choices;
            private string defaultValue;

            private bool hideFromNewForm;
            private bool hideFromEditForm;

            private SPList targetList;
            private string targetFieldName;
            private bool multiLookup;

            private Func<SPFieldCollection, SPField> fieldCreator;
            private bool shouldUpdateAsPartOfCreation = true;

            public FieldDefinition(string name, TParentDefinition parent)
            {
                Name = name;
                this.parent = parent;
                type = SPFieldType.Text;
                shouldRenderAsRichText = false;

                displayName = name;
                IsRequired = true;
                IsHidden = false;

                choices = null;

                hideFromNewForm = false;
                hideFromEditForm = false;

                multiLookup = false;

                fieldCreator = DefaultFieldCreator;
            }

            public FieldDefinition<TParentDefinition> AsDateTime()
            {
                type = SPFieldType.DateTime;
                return this;
            }

            public FieldDefinition<TParentDefinition> AsNumber()
            {
                type = SPFieldType.Number;
                return this;
            }

            public FieldDefinition<TParentDefinition> AsBoolean()
            {
                type = SPFieldType.Boolean;
                return this;
            }

            public FieldDefinition<TParentDefinition> AsSingleUser()
            {
                type = SPFieldType.User;
                return this;
            }

            public FieldDefinition<TParentDefinition> AsCurrency()
            {
                type = SPFieldType.Currency;
                return this;
            }

            public FieldDefinition<TParentDefinition> AsNote()
            {
                type = SPFieldType.Note;
                return this;
            }

            public FieldDefinition<TParentDefinition> AsRichText()
            {
                type = SPFieldType.Note;
                shouldRenderAsRichText = true;
                return this;
            }

            public FieldDefinition<TParentDefinition> AsUrl()
            {
                type = SPFieldType.URL;
                return this;
            }

            public FieldDefinition<TParentDefinition> AsLookupToSelf()
            {
                return AsLookupTo(null, "Title");
            }

            public FieldDefinition<TParentDefinition> AsLookupTo(SPList specificTargetList)
            {
                return AsLookupTo(specificTargetList, "Title");
            }

            public FieldDefinition<TParentDefinition> AsMultiLookupTo(SPList specificTargetList)
            {
                return AsMultiLookupTo(specificTargetList, "Title");
            }

            public FieldDefinition<TParentDefinition> AsLookupTo(SPList specificTargetList, string specificTargetFieldName)
            {
                type = SPFieldType.Lookup;
                targetList = specificTargetList;
                targetFieldName = specificTargetFieldName;
                return this;
            }

            public FieldDefinition<TParentDefinition> AsMultiLookupTo(SPList specificTargetList, string specificTargetFieldName)
            {
                type = SPFieldType.Lookup;
                targetList = specificTargetList;
                targetFieldName = specificTargetFieldName;
                multiLookup = true;
                return this;
            }

            public FieldDefinition<TParentDefinition> WithDisplayName(string title)
            {
                displayName = title;
                return this;
            }

            public FieldDefinition<TParentDefinition> WithInternalName(string internalName)
            {
                Name = internalName;
                return this;
            }

            public FieldDefinition<TParentDefinition> WithChoices(params string[] choiceOptions)
            {
                type = SPFieldType.Choice;
                choices = new StringCollection();
                choices.AddRange(choiceOptions);
                return this;
            }

            public FieldDefinition<TParentDefinition> WithDefaultValue(string value)
            {
                defaultValue = value;
                return this;
            }

            public FieldDefinition<TParentDefinition> HideFromNewForm()
            {
                hideFromNewForm = true;
                return this;
            }

            public FieldDefinition<TParentDefinition> HideFromEditForm()
            {
                hideFromEditForm = true;
                return this;
            }

            public TParentDefinition And
            {
                get { return parent; }
            }

            public TParentDefinition Finally
            {
                get { return parent; }
            }

            public TParentDefinition AsContentTypeDefinition
            {
                get { return parent; }
            }

            public SPField CreateIn(SPFieldCollection targetFields, string groupName)
            {
                var existingField = targetFields.Cast<SPField>().SingleOrDefault(f => f.InternalName.Equals(Name));
                if (existingField != null)
                    return existingField;

                var field = fieldCreator(targetFields);

                if (!string.IsNullOrEmpty(groupName))
                    field.Group = groupName;

                field.Hidden = IsHidden && typeof(TParentDefinition).Equals(typeof(ListDefinition));

                if (!displayName.Equals(Name))
                    field.Title = displayName;

                if (field.Type == SPFieldType.DateTime)
                    ((SPFieldDateTime)field).DisplayFormat = SPDateTimeFieldFormatType.DateOnly;

                if (defaultValue != null)
                    field.DefaultValue = defaultValue;

                if (shouldRenderAsRichText)
                    ((SPFieldMultiLineText)field).RichText = true;

                field.ShowInNewForm = !hideFromNewForm;
                field.ShowInEditForm = !hideFromEditForm;

                if (shouldUpdateAsPartOfCreation)
                    field.Update();

                if (type == SPFieldType.Lookup)
                {
                    var lookup = (SPFieldLookup)field;
                    lookup.LookupField = targetFieldName;
                    lookup.AllowMultipleValues = multiLookup;
                    lookup.Update();
                }

                return field;
            }

            private SPField DefaultFieldCreator(SPFieldCollection targetFields)
            {
                string internalName;
                if (type == SPFieldType.Lookup)
                {
                    if (targetList == null && targetFields.List == null)
                        throw new Exception("You cannot add a self lookup to a content type lookup");
                    internalName = targetFields.AddLookup(displayName, targetList != null ? targetList.ID : targetFields.List.ID, IsRequired);
                }
                else
                    internalName = targetFields.Add(Name, type, IsRequired, false, choices);

                return targetFields.Cast<SPField>().Single(f => f.InternalName.Equals(internalName));
            }
        }

        public class RoleDefinition
        {
            private readonly SPWeb web;
            private readonly string name;

            private string description;
            private SPBasePermissions[] permissions;

            public RoleDefinition(SPWeb web, string name)
            {
                this.web = web;
                this.name = name;

                description = string.Empty;
                permissions = new SPBasePermissions[0];
            }

            public RoleDefinition WithDescription(string newDescription)
            {
                description = newDescription;
                return this;
            }

            public RoleDefinition WithBasePermissions(params SPBasePermissions[] basePermissions)
            {
                permissions = basePermissions;
                return this;
            }

            public void Create()
            {
                var existingDefinition =
                    web.RoleDefinitions.Cast<SPRoleDefinition>().SingleOrDefault(permission => permission.Name.Equals(name));
                if (existingDefinition != null)
                    web.RoleDefinitions.DeleteById(existingDefinition.Id);

                var newDefinition = new SPRoleDefinition { Name = name, Description = description };
                foreach (var permission in permissions)
                    newDefinition.BasePermissions = newDefinition.BasePermissions | permission;
                web.RoleDefinitions.Add(newDefinition);
            }
        }

        public class GroupDefinition
        {
            private readonly SPWeb web;
            private readonly string name;

            private string ownerPrincipalName;
            private string description;

            private bool allowUsersEditMembership;

            public GroupDefinition(SPWeb web, string name)
            {
                this.web = web;
                this.name = name;

                description = string.Empty;
                allowUsersEditMembership = false;
            }

            public GroupDefinition WithOwner(string owner)
            {
                ownerPrincipalName = owner;
                return this;
            }

            public GroupDefinition WithDescription(string newDescription)
            {
                description = newDescription;
                return this;
            }

            public GroupDefinition AllowUsersToEditMembership()
            {
                allowUsersEditMembership = true;
                return this;
            }

            public SPGroup CreateAndReturn()
            {
                var existing = web.SiteGroups.Cast<SPGroup>().SingleOrDefault(g => g.Name.Equals(name));
                if (existing != null)
                    return existing;

                SPPrincipal author = null;
                if (!string.IsNullOrEmpty(ownerPrincipalName))
                    author = web.SiteGroups.Cast<SPGroup>().SingleOrDefault(g => g.Name.Equals(ownerPrincipalName));
                if (author == null)
                    author = web.Users.Cast<SPUser>().SingleOrDefault(u => u.LoginName.Equals(ownerPrincipalName));

                web.SiteGroups.Add(name, author ?? web.Author, null, description);
                var newGroup = web.SiteGroups.Cast<SPGroup>().Single(u => u.Name.Equals(name));

                if (allowUsersEditMembership)
                {
                    newGroup.AllowMembersEditMembership = true;
                    newGroup.Update();
                }

                return newGroup;
            }
        }

        public class RoleAssociation
        {
            private readonly SPWeb web;
            private readonly SPRoleAssignmentCollection roleAssignments;
            private readonly string roleDefinitionName;

            public RoleAssociation(SPWeb web, SPRoleAssignmentCollection roleAssignments, string roleDefinitionName)
            {
                this.web = web;
                this.roleAssignments = roleAssignments;
                this.roleDefinitionName = roleDefinitionName;
            }

            public void For(params string[] principalNames)
            {
                For(principalNames);
            }

            public void For(IEnumerable<string> principalNames)
            {
                foreach (var principal in
                    principalNames.Select(principalName =>
                        (SPPrincipal)web.Users.Cast<SPUser>().SingleOrDefault(u => u.Name.Equals(principalName))
                        ?? web.SiteGroups.Cast<SPGroup>().SingleOrDefault(u => u.Name.Equals(principalName))))
                {
                    For(principal);
                }
            }

            public void For(SPPrincipal principal)
            {
                var assignment = new SPRoleAssignment(principal);
                assignment.RoleDefinitionBindings.Add(web.RoleDefinitions[roleDefinitionName]);

                roleAssignments.Add(assignment);
            }
        }
    }
}