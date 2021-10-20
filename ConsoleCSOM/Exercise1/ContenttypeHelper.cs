using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM.Exercise1
{
    public static class ContenttypeHelper
    {
        public static async Task CreateSiteColumnTypeText(ClientContext ctx)
        {
            ctx.Load(ctx.Web, w => w.Title, c => c.Fields);
            var web = ctx.Web;
            web.Fields.AddFieldAsXml(Constant.FieldAbout, false, AddFieldOptions.AddFieldInternalNameHint);

            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateSiteColumnTypeTaxonomy(ClientContext ctx)
        {
            ctx.Load(ctx.Web, w => w.Title, c => c.Fields);
            var web = ctx.Web;
            // Create as a regular field setting the desired type in XML
            Field field = web.Fields.AddFieldAsXml(Constant.FieldCity, false, AddFieldOptions.AddFieldInternalNameHint);
            ctx.ExecuteQuery();

            Guid termStoreId = Guid.Empty;
            Guid termSetId = Guid.Empty;
            //Get id            
            FieldInfor.GetTaxonomyFieldInfo(ctx, out termStoreId, out termSetId);
            // Retrieve as Taxonomy Field
            TaxonomyField taxonomyField = ctx.CastTo<TaxonomyField>(field);
            taxonomyField.SspId = termStoreId;
            taxonomyField.TermSetId = termSetId;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.AnchorId = Guid.Empty;
            taxonomyField.Update();

            await ctx.ExecuteQueryAsync();
        }

        public static async Task AddFieldToContentType(ClientContext ctx)
        {
            //Get all content type 
            ContentTypeCollection ctp = ctx.Web.ContentTypes;

            //Load
            ctx.Load(ctp);
            ctx.ExecuteQuery();

            //Select contenttype
            ContentType targetContentType = (from contentType in ctp
                                             where contentType.Name == Constant.NameContentType
                                             select contentType).FirstOrDefault();
            ctx.Load(targetContentType);
            ctx.ExecuteQuery();

            Field targetField = ctx.Web.AvailableFields.GetByInternalNameOrTitle(Constant.NameFieldCity);
            FieldLinkCreationInformation flc = new FieldLinkCreationInformation();
            flc.Field = targetField;
            flc.Field.Required = false;
            flc.Field.Hidden = false;

            targetContentType.FieldLinks.Add(flc);

            targetContentType.Update(false);

            await ctx.ExecuteQueryAsync();
        }

        public static async Task AddContenttypeToList(ClientContext ctx)
        {
            ContentTypeCollection ctp = ctx.Web.ContentTypes;

            ctx.Load(ctp);
            ctx.ExecuteQuery();

            ContentType targetContenttype = (from contenttype in ctp
                                             where contenttype.Name == "CSOM Test content type"
                                             select contenttype).FirstOrDefault();

            List targetList = ctx.Web.Lists.GetByTitle("CSOM Test");
            targetList.ContentTypes.AddExistingContentType(targetContenttype);
            targetList.Update();

            ctx.Web.Update();

            await ctx.ExecuteQueryAsync();
        }

        public static async Task SetDefaultContentType(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(Constant.TitleList);
            ContentTypeCollection ctc = list.ContentTypes;

            ctx.Load(ctc);
            ctx.ExecuteQuery();

            IList<ContentTypeId> listId = new List<ContentTypeId>();
            foreach (ContentType ct in ctc)
            {
                if (ct.Name.Equals(Constant.NameContentType))
                {
                    listId.Add(ct.Id);
                }
            }

            list.RootFolder.UniqueContentTypeOrder = listId;
            list.RootFolder.Update();
            list.Update();

            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateListItem(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(Constant.TitleList);

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermGroup termGroup = termStore.Groups.GetByName(Constant.TermGroup);
            TermSet termSet = termGroup.TermSets.GetByName(Constant.TermSet);
            Term term = termSet.Terms.GetByName(Constant.Term_HCM);

            ctx.Load(term);
            ctx.ExecuteQuery();

            ListItem newItem = list.AddItem(itemCreateInfo);
            newItem["Title"] = "5";
            newItem[Constant.FieldInternalNameAbout] = "about 2";
            newItem[Constant.FieldInternalNameCity] = new TaxonomyFieldValue() { TermGuid = term.Id.ToString(), Label = term.Name, WssId = -1 };
            newItem.Update();

            await ctx.ExecuteQueryAsync();
        }

        public static async Task UpdateDefaultValueForFieldAbout(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(Constant.TitleList);
            Field field = list.Fields.GetByInternalNameOrTitle(Constant.FieldInternalNameAbout);
            field.DefaultValue = "about default";
            field.Update();
            await ctx.ExecuteQueryAsync();
        }
        public static async Task UpdateDefaultValueForFieldCity(ClientContext ctx)
        {
            var taxColumn = ctx.CastTo<TaxonomyField>(ctx.Web.Fields.GetByInternalNameOrTitle(Constant.FieldInternalNameCity));
            ctx.Load(taxColumn);
            ctx.ExecuteQuery();
            //get taxonomy field
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermGroup termGroup = termStore.Groups.GetByName(Constant.TermGroup);
            TermSet termSet = termGroup.TermSets.GetByName(Constant.TermSet);
            Term term = termSet.Terms.GetByName(Constant.Term_HCM);

            ctx.Load(term, t => t.Name, t => t.Id);
            ctx.ExecuteQuery();
            //initialize taxonomy field value
            var defaultValue = new TaxonomyFieldValue();
            defaultValue.WssId = -1;
            defaultValue.Label = term.Name;
            defaultValue.TermGuid = term.Id.ToString();
            //retrieve validated taxonomy field value
            var validatedValue = taxColumn.GetValidatedString(defaultValue);
            ctx.ExecuteQuery();
            //set default value for a taxonomy field
            taxColumn.DefaultValue = validatedValue.Value;
            taxColumn.Update();
            await ctx.ExecuteQueryAsync();

        }
    }
}
