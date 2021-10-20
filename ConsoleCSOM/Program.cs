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
using System.Text;
using System.IO;

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

                    //---
                     

                    


                }

                Console.WriteLine($"Press Any Key To Stop!");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
            }
        }        
        
        





        static ClientContext GetContext(ClientContextHelper clientContextHelper)
        {
            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            IConfiguration config = builder.Build();
            //convert json to SharepointInfo
            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();

            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }

      
    }
}
