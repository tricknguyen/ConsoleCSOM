using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM.Exercise1.SharepointServices
{
    public class SharepointService : ISharepointService
    {
        public void CreateTerm(ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            TermGroup termGroup = termStore.Groups.GetByName(Constant.TermGroup);
            TermSet termSet = termGroup.TermSets.GetByName(Constant.TermSet);

            termSet.CreateTerm(Constant.Term_HCM, 1033, Guid.NewGuid());
            termSet.CreateTerm("Stockholm", 1033, Guid.NewGuid());

            ctx.ExecuteQuery();
        }

        public void CreateTermSet(ClientContext ctx)
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(ctx);

            ctx.Load(taxonomySession,
                ts => ts.TermStores.Include(
                    store => store.Name,
                    store => store.Groups.Include(group => group.Name)
                    )
                );

            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            if (termStore != null)
            {
                TermGroup myGroup = termStore.CreateGroup(Constant.TermGroup, Guid.NewGuid());
                //1033 - lcid -locale indetifier for the language
                TermSet myTermSet = myGroup.CreateTermSet(Constant.TermSet, Guid.NewGuid(), 1033);

                ctx.ExecuteQuery();
            }
        }

        public void CreationList(ClientContext ctx)
        {
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = Constant.TitleList;
            creationInfo.Description = "New List created by VN";
            creationInfo.TemplateType = (int)ListTemplateType.GenericList;

            List newList = ctx.Web.Lists.Add(creationInfo);

            ctx.Load(newList);
            ctx.ExecuteQuery();
        }

        public void CreateSiteContentType(ClientContext ctx)
        {
            //CREATE CONTENT TYPE
            // Get the content type collection for the website
            ContentTypeCollection contentTypeColl = ctx.Web.ContentTypes;

            // Specifies properties that are used as parameters to initialize a new content type.
            ContentTypeCreationInformation newCt = new ContentTypeCreationInformation();
            newCt.Name = Constant.NameContentType;
            newCt.Description = "Training";
            newCt.Group = "List Content Types";

            // Add the new content type to the collection
            ContentType ct = contentTypeColl.Add(newCt);
            ctx.Load(ct);
            ctx.ExecuteQuery();
        }
    }
}
