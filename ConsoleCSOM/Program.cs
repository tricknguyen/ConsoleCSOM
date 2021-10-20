using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;
using Microsoft.SharePoint.Client.UserProfiles;
using ConsoleCSOM.Exercise1.SharepointServices;

namespace ConsoleCSOM
{
    class Program
    {
        private static ISharepointService _services;
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper()) 
                {
                    ClientContext ctx = GetContext(clientContextHelper);

                    //Get all content type 
                    ContentTypeCollection ctp = ctx.Web.ContentTypes;

                    //Load
                    ctx.Load(ctp);
                    ctx.ExecuteQuery();

                    //Select contenttype
                    ContentType targetContentType = (from contentType in ctp
                                                     where contentType.Name == "CSOM Test content type"
                                                     select contentType).FirstOrDefault();
                    ctx.Load(targetContentType);
                    ctx.ExecuteQuery();

                    Field targetField = ctx.Web.AvailableFields.GetByInternalNameOrTitle("cities");
                    FieldLinkCreationInformation flc = new FieldLinkCreationInformation();
                    flc.Field = targetField;
                    flc.Field.Required = false;
                    flc.Field.Hidden = false;

                    targetContentType.FieldLinks.Add(flc);
                    targetContentType.Update(false);

                    List targetList = ctx.Web.Lists.GetByTitle("CSOM Test");
                    //targetList.ContentTypes.AddExistingContentType(targetContentType);
                    
                    
                   
                    targetList.Update();

                    await ctx.ExecuteQueryAsync();








                }

                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
            }
        }        
        //exercise 3
        private static async Task CreateTaxonomyField (ClientContext ctx)
        {
            ctx.Load(ctx.Web, w => w.Title, c => c.Fields);
            var web = ctx.Web;
            // Create as a regular field setting the desired type in XML
            Field field = web.Fields.AddFieldAsXml("<Field DisplayName='cities' " +
                "Name='cities' " +
                "ID='{159B9BB9-982C-4A8B-96C8-8D7031F1DE80}' " +
                "Group='Custom Columns' Type='TaxonomyFieldTypeMulti' />", false, AddFieldOptions.AddFieldInternalNameHint);
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

            taxonomyField.AllowMultipleValues = true;

            taxonomyField.Update();

            await ctx.ExecuteQueryAsync();
        }


        



        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            //convert json to SharepointInfo
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();

            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

        private static async Task GetFieldTermValue(ClientContext Ctx, string termId)
        {
            //load term by id
            TaxonomySession session = TaxonomySession.GetTaxonomySession(Ctx);
          
            Term taxonomyTerm = session.GetTerm(new Guid(termId));
            Ctx.Load(taxonomyTerm, t => t.Labels,
                                   t => t.Name,
                                   t => t.Id);
            await Ctx.ExecuteQueryAsync();
        }

        private static async Task ExampleSetTaxonomyFieldValue(ListItem item, ClientContext ctx)
        {
            var field = ctx.Web.Fields.GetByTitle("fieldname");

            ctx.Load(field);
            await ctx.ExecuteQueryAsync();

            var taxField = ctx.CastTo<TaxonomyField>(field);

            taxField.SetFieldValueByValue(item, new TaxonomyFieldValue()
            {
                WssId = -1, // alway let it -1
                Label = "correct label here",
                TermGuid = "term id"
            });
            item.Update();
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomTermSetAsync(ClientContext ctx)
        {
            // Get the TaxonomySession
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            // Get the term store by name
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName("Test");
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName("Test Term Set");


            var terms = termSet.GetAllTerms();

            ctx.Load(terms);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CsomLinqAsync(ClientContext ctx)
        {
            var fieldsQuery = from f in ctx.Web.Fields
                              where f.InternalName == "Test" ||
                                    f.TypeAsString == "TaxonomyFieldTypeMulti" ||
                                    f.TypeAsString == "TaxonomyFieldType"
                              select f;

            var fields = ctx.LoadQuery(fieldsQuery);
            await ctx.ExecuteQueryAsync();
        }

        private static async Task SimpleCamlQueryAsync(ClientContext ctx)
        {
            var list = ctx.Web.Lists.GetByTitle("Documents");

            var allItemsQuery = CamlQuery.CreateAllItemsQuery();
            var allFoldersQuery = CamlQuery.CreateAllFoldersQuery();

            var items = list.GetItems(new CamlQuery()
            {
                ViewXml = @"<View>
                                <Query>
                                    <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>
                                </Query>
                                <RowLimit>20</RowLimit>
                            </View>",
                FolderServerRelativeUrl = "/sites/test-site-duc-11111/Shared%20Documents/2"
                //example for site: https://omniapreprod.sharepoint.com/sites/test-site-duc-11111/
            });

            ctx.Load(items);
            await ctx.ExecuteQueryAsync();
        }
    }
}
