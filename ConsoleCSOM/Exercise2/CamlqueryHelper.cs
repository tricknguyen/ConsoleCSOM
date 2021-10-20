using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM.Exercise2
{
    public class CamlqueryHelper
    {
        public static async Task QueryListItem(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(Constant.TitleList);
            ctx.Load(list.ContentTypes);
            await ctx.ExecuteQueryAsync();

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = QueryCaml.QueryListItem;

            ListItemCollection cli = list.GetItems(camlQuery);

            ctx.Load(cli);
            ctx.ExecuteQuery();

            var item = cli.FirstOrDefault();
            Console.WriteLine(item.DisplayName);
        }

        private static async Task CreateListView(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(Constant.TitleList);

            ViewCollection viewCollection = list.Views;
            ctx.Load(viewCollection);
            ctx.ExecuteQuery();

            ViewCreationInformation viewCreationInformation = new ViewCreationInformation();

            viewCreationInformation.Title = Constant.TitleListView;
            viewCreationInformation.ViewTypeKind = ViewType.None;
            viewCreationInformation.RowLimit = 10;

            viewCreationInformation.Query = QueryCaml.QueryListView;

            string CommaSeparateColumnNames = "ID,Test_x0020_2,City_x0020_Hunter";
            viewCreationInformation.ViewFields = CommaSeparateColumnNames.Split(',');

            viewCollection.Add(viewCreationInformation);

            await ctx.ExecuteQueryAsync();
        }

        private static async Task UpdateListItem(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(Constant.TitleList);

            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = QueryCaml.QueryUpdateListItem;

            ListItemCollection listItems = list.GetItems(camlQuery);
            ctx.Load(listItems);
            ctx.ExecuteQuery();

            foreach (var item in listItems)
            {
                item[Constant.FieldInternalNameAbout] = "Update script";
                item.Update();
            }
            await ctx.ExecuteQueryAsync();
        }

        private static async Task CreateFieldAuthor(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(Constant.TitleList);
            Field field = list.Fields.AddFieldAsXml(Constant.FieldAuthor, true, AddFieldOptions.AddFieldInternalNameHint);
            ctx.Load(field);
            await ctx.ExecuteQueryAsync();
        }

        public static async Task SetAdminToAuthor(ClientContext ctx)
        {
            List list = ctx.Web.Lists.GetByTitle(Constant.TitleList);

            CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery();

            ListItemCollection listItems = list.GetItems(camlQuery);


            ctx.Load(listItems);
            ctx.ExecuteQuery();


            foreach (var item in listItems)
            {
                item["author0"] = ctx.Web.CurrentUser;
                item.Update();
            }
            await ctx.ExecuteQueryAsync();
        }
    }
}
