using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Collections.Generic;

namespace ConsoleCSOM
{
    class SharepointInfo
    {
        public string SiteUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                using (var clientContextHelper = new ClientContextHelper()) 
                {
                    ClientContext ctx = GetContext(clientContextHelper);
                    
                }

                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
            }
        }
        private static async Task CreationList(ClientContext ctx)
        {
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = "CSOM Test";
            creationInfo.Description = "New List created by VN";
            creationInfo.TemplateType = (int)ListTemplateType.GenericList;

            List newList = ctx.Web.Lists.Add(creationInfo);

            ctx.Load(newList);
            await ctx.ExecuteQueryAsync();

        }

        public static async Task CreateTermSet(ClientContext ctx)
        {
            string termGroupName = "NewTermGroup";
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);

            ctx.Load(taxonomySession, 
                ts => ts.TermStores.Include(
                    store => store.Name,
                    store => store.Groups.Include(group => group.Name)
                    )
                );

            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            if (termStore!=null)
            {
                TermGroup myGroup = termStore.CreateGroup(termGroupName, Guid.NewGuid());
                //1033 - lcid -locale indetifier for the language
                TermSet myTermSet = myGroup.CreateTermSet("city-anhvu", Guid.NewGuid(), 1033);

                await ctx.ExecuteQueryAsync();
            }
            
        }

        public static async Task CreateTerm(ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermGroup termGroup = termStore.Groups.GetByName("NewTermGroup");
            TermSet termSet = termGroup.TermSets.GetByName("city-anhvu");

            termSet.CreateTerm("Ho Chi Minh", 1033, Guid.NewGuid());
            termSet.CreateTerm("Stockholm", 1033, Guid.NewGuid());

            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateSiteColumnTypeText(ClientContext ctx)
        {
            ctx.Load(ctx.Web, w => w.Title, c => c.Fields);
            var web = ctx.Web;
            web.Fields.AddFieldAsXml("<Field DisplayName='Test 2' " +
                "Name='Test 2' " +
                "ID='{15BB3A47-ABD4-4ED9-9636-51791B0DB550}' " +
                "Group='Custom Columns' " +
                "Type='Text' />", false, AddFieldOptions.AddFieldInternalNameHint);
            
            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateSiteColumnTypeTaxonomy(ClientContext ctx)
        {
            ctx.Load(ctx.Web, w => w.Title, c => c.Fields);
            var web = ctx.Web;
            // Create as a regular field setting the desired type in XML
            Field field = web.Fields.AddFieldAsXml("<Field DisplayName='City Hunter' " +
                "Name='City Hunter' " +
                "ID='{98FBA25E-625C-46EC-887F-587B4072CFD4}' " +
                "Group='Custom Columns' Type='TaxonomyFieldTypeMulti' />", false, AddFieldOptions.AddFieldInternalNameHint);
            ctx.ExecuteQuery();

            Guid termStoreId = Guid.Empty;
            Guid termSetId = Guid.Empty;
            //Get id
            GetTaxonomyFieldInfo(ctx, out termStoreId, out termSetId);

            // Retrieve as Taxonomy Field
            TaxonomyField taxonomyField = ctx.CastTo<TaxonomyField>(field);
            taxonomyField.SspId = termStoreId;
            taxonomyField.TermSetId = termSetId;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.AnchorId = Guid.Empty;
            taxonomyField.Update();

            await ctx.ExecuteQueryAsync();
        }

        private static void GetTaxonomyFieldInfo(ClientContext clientContext, out Guid termStoreId, out Guid termSetId)
        {
            termStoreId = Guid.Empty;
            termSetId = Guid.Empty;

            TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
            TermStore termStore = session.GetDefaultSiteCollectionTermStore();
            TermSetCollection termSets = termStore.GetTermSetsByName("city-anhvu", 1033);

            clientContext.Load(termSets, tsc => tsc.Include(ts => ts.Id));
            clientContext.Load(termStore, ts => ts.Id);
            clientContext.ExecuteQuery();

            termStoreId = termStore.Id;
            termSetId = termSets.FirstOrDefault().Id;
        }

        public static async Task CreateSiteContentType (ClientContext ctx)
        {
            //CREATE CONTENT TYPE
            // Get the content type collection for the website
            ContentTypeCollection contentTypeColl = ctx.Web.ContentTypes;

            // Specifies properties that are used as parameters to initialize a new content type.
            ContentTypeCreationInformation newCt = new ContentTypeCreationInformation();
            newCt.Name = "CSOM Test content type";
            newCt.Description = "Training";
            newCt.Group = "List Content Types";

            // Add the new content type to the collection
            ContentType ct = contentTypeColl.Add(newCt);
            ctx.Load(ct);
            await ctx.ExecuteQueryAsync();

            // Display that the content type is created.
            Console.WriteLine(ct.Name + " content type is created successfully");
        }

        public static async Task AddFieldToContentType (ClientContext ctx)
        {
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

            Field targetField = ctx.Web.AvailableFields.GetByInternalNameOrTitle("City Hunter");
            FieldLinkCreationInformation flc = new FieldLinkCreationInformation();
            flc.Field = targetField;
            flc.Field.Required = false;
            flc.Field.Hidden = false;

            targetContentType.FieldLinks.Add(flc);
            targetContentType.Update(false);

            await ctx.ExecuteQueryAsync();
        }

        public static async Task AddContenttypeToList (ClientContext ctx)
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

        public static async Task SetDefaultContentType (ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");
            ContentTypeCollection ctc = list.ContentTypes;

            ctx.Load(ctc);
            ctx.ExecuteQuery();

            IList<ContentTypeId> listId = new List<ContentTypeId>();
            foreach (ContentType ct in ctc)
            {
                if (ct.Name.Equals("CSOM Test content type"))
                {
                    listId.Add(ct.Id);
                }
            }

            list.RootFolder.UniqueContentTypeOrder = listId;
            list.RootFolder.Update();
            list.Update();

            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateListItem (ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermGroup termGroup = termStore.Groups.GetByName("NewTermGroup");
            TermSet termSet = termGroup.TermSets.GetByName("city-anhvu");
            Term term = termSet.Terms.GetByName("Ho Chi Minh");

            ctx.Load(term);
            ctx.ExecuteQuery();

            ListItem newItem = list.AddItem(itemCreateInfo);
            newItem["Title"] = "5";
            newItem["Test_x0020_2"] = "about 2";
            newItem["City_x0020_Hunter"] = new TaxonomyFieldValue() { TermGuid = term.Id.ToString(), Label = term.Name, WssId = -1 };
            newItem.Update();

            await ctx.ExecuteQueryAsync();
        }

        public static async Task UpdateDefaultValueForFieldAbout (ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");
            Field field = list.Fields.GetByInternalNameOrTitle("Test_x0020_2");
            field.DefaultValue = "about default";
            field.Update();
            await ctx.ExecuteQueryAsync();
        }
        public static async Task UpdateDefaultValueForFieldCity (ClientContext ctx)
        {
            var taxColumn = ctx.CastTo<TaxonomyField>(ctx.Web.Fields.GetByInternalNameOrTitle("City_x0020_Hunter"));
            ctx.Load(taxColumn);
            ctx.ExecuteQuery();
            //get taxonomy field
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermGroup termGroup = termStore.Groups.GetByName("NewTermGroup");
            TermSet termSet = termGroup.TermSets.GetByName("city-anhvu");
            Term term = termSet.Terms.GetByName("Ho Chi Minh");

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

        private static async Task QueryListItem(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle("CSOM Test");

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"<View>
                                <Query>
                                    <Where>
                                        <Neq>
                                            <FieldRef Name='Test'/>
                                            <Value Type='Text'>about default</Value>
                                        </Neq>
                                    </Where>
                                </Query>
                                <RowLimit>100</RowLimit>
                            </View>";
            ListItemCollection cli = list.GetItems(camlQuery);

            ctx.ExecuteQuery();

            //foreach (ListItem item in cli)
            //{
            //    Console.WriteLine(item.Tes);
            //}

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
