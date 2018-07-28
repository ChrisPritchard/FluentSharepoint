using System;
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;

namespace FluentSharePoint
{
    public static partial class FluentCreationExtensions
    {
        public partial interface ICanHaveFields
        {
            string TaxonomyTermStoreName { get; set; }
            string TaxonomyGroupName { get; set; }
        }

        public partial class ContentTypeDefinition
        {
            public string TaxonomyTermStoreName { get; set; }
            public string TaxonomyGroupName { get; set; }
            
            public ContentTypeDefinition UsingTaxonomyTermStore(string termStoreName)
            {
                TaxonomyTermStoreName = termStoreName;
                return this;
            }

            public ContentTypeDefinition UsingTaxonomyGroup(string groupName)
            {
                TaxonomyGroupName = groupName;
                return this;
            }
        }

        public partial class ListDefinition
        {
            public string TaxonomyTermStoreName { get; set; }
            public string TaxonomyGroupName { get; set; }
            
            public ListDefinition UsingTaxonomyTermStore(string termStoreName)
            {
                TaxonomyTermStoreName = termStoreName;
                return this;
            }

            public ListDefinition UsingTaxonomyGroup(string groupName)
            {
                TaxonomyGroupName = groupName;
                return this;
            }
        }

        public partial class FieldDefinition<TParentDefinition>
        {
            const string defaultTermStore = "Managed Metadata Service";

            enum TaxonomyType { None, SingleOpen, SingleClosed, MultipleOpen, MultipleClosed }

            private TaxonomyType taxonomyType = TaxonomyType.None;
            private string taxonomyTermSetName;

            public FieldDefinition<TParentDefinition> AsSingleOpenTaxonomyAgainst(string termSetName)
            {
                ConfigureAsTaxonomy(TaxonomyType.SingleOpen, termSetName);
                return this;
            }

            public FieldDefinition<TParentDefinition> AsSingleClosedTaxonomyAgainst(string termSetName)
            {
                ConfigureAsTaxonomy(TaxonomyType.SingleClosed, termSetName);
                return this;
            }

            public FieldDefinition<TParentDefinition> AsMultipleOpenTaxonomyAgainst(string termSetName)
            {
                ConfigureAsTaxonomy(TaxonomyType.MultipleOpen, termSetName);
                return this;
            }

            public FieldDefinition<TParentDefinition> AsMultipleClosedTaxonomyAgainst(string termSetName)
            {
                ConfigureAsTaxonomy(TaxonomyType.MultipleClosed, termSetName);
                return this;
            }

            private void ConfigureAsTaxonomy(TaxonomyType specificType, string termSetName)
            {
                taxonomyType = specificType;
                taxonomyTermSetName = termSetName;
                fieldCreator = CreateTaxonomyField;
                shouldUpdateAsPartOfCreation = false;
            }

            private SPField CreateTaxonomyField(SPFieldCollection targetFields)
            {
                if(string.IsNullOrEmpty(parent.TaxonomyGroupName))
                    throw new Exception("The current definition does not have a TaxonomyGroupName set");

                var session = new TaxonomySession(parent.Web.Site);
                var termStore = session.TermStores.Single(s => s.Name.Equals(parent.TaxonomyTermStoreName ?? defaultTermStore));
                var group = termStore.Groups.Single(g => g.Name.Equals(parent.TaxonomyGroupName));

                var createName = !displayName.Equals(Name) ? displayName : Name;
                var field = (TaxonomyField)targetFields.CreateNewField("TaxonomyFieldType", createName);
                field.AllowMultipleValues = taxonomyType == TaxonomyType.MultipleOpen || taxonomyType == TaxonomyType.MultipleClosed;
                field.SspId = termStore.Id;
                field.TermSetId = group.TermSets.Single(t => t.Name.Equals(taxonomyTermSetName)).Id;
                field.CreateValuesInEditForm = taxonomyType == TaxonomyType.MultipleOpen || taxonomyType == TaxonomyType.SingleOpen;

                targetFields.Add(field);
                return targetFields[createName];
            }
        }
    }
}