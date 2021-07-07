using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Concurrent;
using System.Threading.Tasks;

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
                    ctx.Load(ctx.Web);
                    await ctx.ExecuteQueryAsync();

                    Console.WriteLine($"Site {ctx.Web.Title}");
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

            var info = config.GetSection("SharepointInfo").Get<SharepointInfo>();

            return clientContextHelper.GetContext(new Uri(info.SiteUrl), info.Username, info.Password);
        }
    }
}
