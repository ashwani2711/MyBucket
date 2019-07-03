using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SP.GetWebPartUsage
{
    class Program
    {
        #region Variables
        static string sUrl = null;
        static string username = null;
        static string domain = null;
        static bool IsSPOnline = false;
        static FileStream ostrm;
        static StreamWriter writer;
        static TextWriter oldOut = Console.Out;
        static string WebPartsToRetrieve = null;
        static bool GetRootWeb = false;
        static string WebNames = null;
       internal static string strOutput = null;
        #endregion

        static void Main(string[] args)
        {
            #region Get Configuration
            try
            {
                //sUrl = ConfigurationManager.AppSettings["siteurl"].ToString();
                username = ConfigurationManager.AppSettings["username"].ToString();
                domain = ConfigurationManager.AppSettings["domain"].ToString();
                IsSPOnline = bool.Parse(ConfigurationManager.AppSettings["IsSPOnline"].ToString());
                WebPartsToRetrieve = ConfigurationManager.AppSettings["WebPartsToRetrieve"].ToString();
                //WebNames = ConfigurationManager.AppSettings["WebName"].ToString();
                //GetRootWeb = bool.Parse(ConfigurationManager.AppSettings["GetRootWeb"].ToString());
            }
            catch { }
            #endregion

            #region Configure Output
            DateTime startTime;
            DateTime endTime;
            //Initialize the Console Output (to file and also to console)
            try
            {
                ostrm = new FileStream(Environment.CurrentDirectory + "\\OutputLog - " +
                            DateTime.Now.ToString("dd-MM-yyyy hh mm ss") + ".txt", FileMode.OpenOrCreate, FileAccess.Write);
                writer = new StreamWriter(ostrm);
            }
            catch (Exception e)
            {
                WriteLine("Cannot open Redirect.txt for writing");
                WriteLine(e.Message);
                return;
            }

            //Set the default output to Console
            Console.SetOut(oldOut);
            #endregion
            try
            {
                //If the required parameters are provided
                if (!string.IsNullOrEmpty(username))
                {
                    SecureString pwd = GetPasswordFromText();
                    startTime = DateTime.Now;
                    WriteLine("Starting the Report at " + startTime.ToString("dd-MM-yyyy HH:mm:ss"));
                    //Get the Secured Password
                    List<string> sitesList = ReadCsvFile();
                    Program.strOutput += "PageUrl,StorageKey,WebPartID,WebPartTitle,WebPartType,WebURL,ZoneID,ZoneIndex" + Environment.NewLine;
                    foreach (string st in sitesList)
                    {
                        //Initialize teh File Status object with required parameters
                        WebPartUsage webPartUsage = new WebPartUsage(
                            st,
                            username,
                            domain,
                            IsSPOnline,
                            WebPartsToRetrieve,
                            true,
                            null,
                            writer,
                            oldOut, pwd);

                        //Generate the Reports for the Site Collection
                        webPartUsage.ProcessSiteCollection();
                    }

                    if (Program.strOutput != "Page Url\tWebPartId\tWeb Part Title\tWeb Part Type\tWebURL\tZoneIndex" + Environment.NewLine)
                    {
                        System.IO.File.WriteAllText(Environment.CurrentDirectory + "\\WebPartUsage-" + DateTime.Now.ToString("dd-MM-yyyy HH mm ss") + ".txt", Program.strOutput);
                    }
                    endTime = DateTime.Now;

                    WriteLine("Operation Completed Successfully");
                    WriteLine("Completion time: " + endTime.ToString("dd-MM-yyyy HH:mm:ss"));
                    TimeSpan t = new TimeSpan();
                    t = endTime.Subtract(startTime);
                    WriteLine("Execution Time: " + t.Days + " days " +
                                                    t.Hours + " hours " +
                                                    t.Minutes + " minutes " +
                                                    t.Seconds + " seconds");
                }
                else
                {
                    WriteLine("Configuration information not found");
                    WriteLine("Please check the configuration (.config) and run the script again");
                }

            }
            catch (Exception ex)
            {
                WriteLine(ex.Message);
            }
            finally
            {
                writer.Close();
                ostrm.Close();
            }
            Console.ReadLine();
        }

        /// <summary>
        /// Writes the Output to Console and also to the file
        /// </summary>
        /// <param name="message">Message to be written to the file and the console</param>
        static void WriteLine(string message)
        {
            Console.SetOut(writer);
            Console.WriteLine(message);
            Console.SetOut(oldOut);
            Console.WriteLine(message);
        }

        static List<string> ReadCsvFile()
        {
            var reader = new StreamReader(System.IO.File.OpenRead(@"siteCollections.csv"));
            List<string> listSites = new List<string>();
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (!string.IsNullOrEmpty(line))
                    listSites.Add(line.Trim());
            }
            return listSites;
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
            WriteLine(Environment.NewLine);
            return pwd;
        }
        #endregion
    }
}
