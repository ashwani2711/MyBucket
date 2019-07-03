using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SP.EmptyRecycleBin
{
    class Program
    {
        static ClientContext context = null;
        static void Main(string[] args)
        {
            string strUrl = ConfigurationManager.AppSettings["siteurl"].ToString();
            string username = ConfigurationManager.AppSettings["username"].ToString();
            bool isonline = bool.Parse(ConfigurationManager.AppSettings["isonline"].ToString());
            string domain = ConfigurationManager.AppSettings["domain"].ToString();

            try
            {
                context = new ClientContext(strUrl);
                context.RequestTimeout = 3600000;
                if (isonline)
                {
                    context.Credentials = new SharePointOnlineCredentials(username, GetPasswordFromText());
                }
                else
                {
                    context.Credentials = new NetworkCredential(username, GetPasswordFromText(), domain);
                }

                Site site = context.Site;
                //Web rootWeb = site.RootWeb;

                context.Load(site);
                context.ExecuteQuery();

                Console.WriteLine("Emptying the Recycle Bin");
                site.RecycleBin.DeleteAll();
                context.ExecuteQuery();
                Console.WriteLine("Recyle Bin emptied");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                context.Dispose();
            }

            Console.WriteLine("Operation Completed Successfully");
            Console.ReadLine();
        }

        private static void IterateSites(Web web)
        {
            context.Load(web, w => w.RecycleBin, w => w.Webs, w => w.Url);
            context.ExecuteQuery();

            Console.WriteLine("Emptying the recycle bin on: " + web.Url);
            web.RecycleBin.DeleteAll();
            context.ExecuteQuery();
            Console.WriteLine("Completed: " + web.Url + Environment.NewLine);

            foreach(Web subWeb in web.Webs)
            {
                IterateSites(subWeb);
            }
        }


        #region Get Password From Text
        /// <summary>
        /// Converts the normal string to SecuredString
        /// </summary>
        /// <param name="password"></param>
        /// <returns></returns>
        public static SecureString GetPasswordFromText()
        {
            Console.Write("Enter your password: ");
            SecureString pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }
            Console.WriteLine(Environment.NewLine);
            return pwd;
        }
        #endregion
    }
}
