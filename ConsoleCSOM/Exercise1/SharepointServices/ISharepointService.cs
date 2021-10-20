using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM.Exercise1.SharepointServices
{
    public interface ISharepointService
    {
        void CreationList(ClientContext ctx);
        void CreateTermSet(ClientContext ctx);
        void CreateTerm(ClientContext ctx);
        void CreateSiteContentType(ClientContext ctx);
    }
}
