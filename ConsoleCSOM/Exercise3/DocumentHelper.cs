using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM.Exercise3
{
    public static class DocumentHelper
    {
        public static async Task CreateTaxonomyField(ClientContext ctx)
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

        public static async Task AddfieldCitiesToCT(ClientContext ctx)
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

            Field targetField = ctx.Web.AvailableFields.GetByInternalNameOrTitle("cities");
            FieldLinkCreationInformation flc = new FieldLinkCreationInformation();
            flc.Field = targetField;
            flc.Field.Required = false;
            flc.Field.Hidden = false;

            targetContentType.FieldLinks.Add(flc);
            targetContentType.Update(true);

            await ctx.ExecuteQueryAsync();
        }

        public static async Task AddListItem(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(Constant.TitleList);

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

            TermSet termSet = GetTermSet(ctx);
            Term term1 = termSet.Terms.GetByName(Constant.Term_HCM);
            Term term2 = termSet.Terms.GetByName("Stockholm");

            Field field = ctx.Web.AvailableFields.GetByInternalNameOrTitle("cities");

            ctx.Load(ctx.Web.CurrentUser);
            ctx.Load(term1);
            ctx.Load(term2);
            ctx.Load(field);
            ctx.ExecuteQuery();

            TaxonomyField txfield = ctx.CastTo<TaxonomyField>(field);
            TaxonomyFieldValueCollection termValues = null;
            string termValueString = $"-1;#{term1.Name}|{term1.Id.ToString()};#-1;#{term2.Name}|{term2.Id.ToString()}";
            termValues = new TaxonomyFieldValueCollection(ctx, termValueString, txfield);

            ListItem newItem = list.AddItem(itemCreateInfo);
            newItem["Title"] = "Exercis-3-3-3";
            newItem[Constant.FieldInternalNameAbout] = "about12323";
            newItem[Constant.FieldInternalNameCity] = new TaxonomyFieldValue() { TermGuid = term1.Id.ToString(), Label = term1.Name, WssId = -1 };
            txfield.SetFieldValueByValueCollection(newItem, termValues);
            newItem.Update();

            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateDocumentList(ClientContext ctx)
        {
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = Constant.TitleDocument;
            creationInfo.Description = "Document Test";
            creationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary;

            List newList = ctx.Web.Lists.Add(creationInfo);

            ctx.Load(newList);
            await ctx.ExecuteQueryAsync();
        }

        public static async Task AddCTToDocumentList(ClientContext ctx)
        {
            ContentTypeCollection ctp = ctx.Web.ContentTypes;

            ctx.Load(ctp);
            ctx.ExecuteQuery();

            ContentType targetContenttype = (from contenttype in ctp
                                             where contenttype.Name == Constant.NameContentType
                                             select contenttype).FirstOrDefault();

            List targetDList = ctx.Web.Lists.GetByTitle(Constant.TitleDocument);
            targetDList.ContentTypes.AddExistingContentType(targetContenttype);
            targetDList.Update();

            ctx.Web.Update();

            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateFolder(ClientContext ctx)
        {
            //Enable foldercreation for list
            List targetDList = ctx.Web.Lists.GetByTitle(Constant.TitleDocument);
            targetDList.EnableFolderCreation = true;
            targetDList.Update();
            ctx.ExecuteQuery();

            //create folder
            ListItemCreationInformation itemCreationInfo = new ListItemCreationInformation();
            itemCreationInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
            itemCreationInfo.LeafName = "Folder 1";

            ListItem newItem = targetDList.AddItem(itemCreationInfo);
            newItem["Title"] = "Folder 1";
            newItem.Update();

            ctx.ExecuteQuery();
        }

        public static async Task CreateSubFolder(ClientContext ctx)
        {
            ctx.Load(ctx.Web, c => c.ServerRelativeUrl);
            ctx.ExecuteQuery();

            Folder folder1 = ctx.Web.GetFolderByServerRelativeUrl(ctx.Web.ServerRelativeUrl + "/Document%20Test/Folder%201");

            folder1.Folders.Add("Folder 2");
            folder1.Update();

            await ctx.ExecuteQueryAsync();
        }

        public static async Task CreateFileInFolder(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(Constant.TitleDocument);

            /*TermSet termSet = GetTermSet(ctx);
            Term term = termSet.Terms.GetByName("Stockholm");

            Field field = ctx.Web.AvailableFields.GetByInternalNameOrTitle("cities");
            TaxonomyField txfield = ctx.CastTo<TaxonomyField>(field);*/

            //ctx.Load(field);
            ctx.Load(ctx.Web, c => c.ServerRelativeUrl);
            //ctx.Load(term);
            ctx.ExecuteQuery();


            /*TaxonomyFieldValue value = new TaxonomyFieldValue();
            value.Label = term.Name;
            value.TermGuid = term.Id.ToString();
            value.WssId = -1;*/

            FileCreationInformation createFile = new FileCreationInformation();
            createFile.Url = "test4.txt";
            //use byte array to set content of the file
            string somestring = "hello there";
            byte[] toBytes = Encoding.ASCII.GetBytes(somestring);

            createFile.Content = toBytes;

            Folder folder2 = ctx.Web.GetFolderByServerRelativeUrl(ctx.Web.ServerRelativeUrl + "/Document%20Test/Folder%201/Folder%202");
            Microsoft.SharePoint.Client.File addedFile = folder2.Files.Add(createFile);
            ctx.Load(addedFile);
            ctx.ExecuteQuery();

            ListItem item = addedFile.ListItemAllFields;
            item["Title"] = "4";
            item[Constant.FieldInternalNameAbout] = "Folder test";
            //txfield.SetFieldValueByValue(item, value);
            item.Update();
            ctx.Load(item);
            ctx.ExecuteQuery();
        }

        public static async Task QueryItemInFolder(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(Constant.TitleDocument);

            ctx.Load(ctx.Web, c => c.ServerRelativeUrl);
            ctx.ExecuteQuery();

            Folder folder2 = ctx.Web.GetFolderByServerRelativeUrl(ctx.Web.ServerRelativeUrl + "/Document%20Test/Folder%201/Folder%202");
            ctx.Load(folder2);
            ctx.ExecuteQuery();

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = QueryCaml.QueryItemInFolder;
            camlQuery.FolderServerRelativeUrl = folder2.ServerRelativeUrl;

            ListItemCollection cli = list.GetItems(camlQuery);

            ctx.Load(cli);
            ctx.ExecuteQuery();

            var item = cli.FirstOrDefault();
            Console.WriteLine(item["Title"]);
        }

        public static async Task CreateListItemByUpload(ClientContext ctx)
        {
            ctx.Load(ctx.Web, c => c.ServerRelativeUrl);
            ctx.ExecuteQuery();

            string uploadFolderUrl = ctx.Web.ServerRelativeUrl + "/Document%20Test";

            var fileCreation = new FileCreationInformation
            {
                Content = System.IO.File.ReadAllBytes(Constant.uploadFilePath),
                Overwrite = true,
                Url = Path.GetFileName(Constant.uploadFilePath)
            };
            var targetFolder = ctx.Web.GetFolderByServerRelativeUrl(uploadFolderUrl);
            var uploadFile = targetFolder.Files.Add(fileCreation);
            ctx.Load(uploadFile);
            await ctx.ExecuteQueryAsync();
        }

        public static TermSet GetTermSet(ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermGroup termGroup = termStore.Groups.GetByName(Constant.TermGroup);
            TermSet termSet = termGroup.TermSets.GetByName(Constant.TermSet);
            return termSet;
        }

    }
}
