using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SP.GetCheckOutFileDetails
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
        static bool GetRootWeb = false;
        static string WebNames = null;
        #endregion

        /// <summary>
        /// Entry point
        /// </summary>
        /// <param name="args">Arguments</param>
        static void Main(string[] args)
        {
            bool UndoCheckOut = false;
            bool DeleteIfNoPrevVersions = false;

            #region Get Configuration
            try
            {
                sUrl = ConfigurationManager.AppSettings["siteurl"].ToString();
                username = ConfigurationManager.AppSettings["username"].ToString();
                domain = ConfigurationManager.AppSettings["domain"].ToString();
                IsSPOnline = bool.Parse(ConfigurationManager.AppSettings["IsSPOnline"].ToString());
                UndoCheckOut = bool.Parse(ConfigurationManager.AppSettings["UndoCheckOut"].ToString());
                DeleteIfNoPrevVersions = bool.Parse(ConfigurationManager.AppSettings["DeleteIfNoPrevVersions"].ToString());
                WebNames = ConfigurationManager.AppSettings["WebName"].ToString();
                GetRootWeb = bool.Parse(ConfigurationManager.AppSettings["GetRootWeb"].ToString());
            }
            catch { }
            #endregion

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

            try
            {
                //If the required parameters are provided
                if (!string.IsNullOrEmpty(sUrl) && !string.IsNullOrEmpty(username))
                {
                    startTime = DateTime.Now;
                    WriteLine("Starting the Report at " + startTime.ToString("dd-MM-yyyy HH:mm:ss"));
                    //Get the Secured Password

                    if (UndoCheckOut)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.Write("Do you want to perform undo check out operation? " + Environment.NewLine +
                                    "Note that it will undo the checkout and users changes will be lost? (y/n): ");
                        Console.ResetColor();

                        if (Console.ReadLine().Equals("n", StringComparison.InvariantCultureIgnoreCase))
                        {
                            UndoCheckOut = false;
                        }
                    }
                    //Initialize teh File Status object with required parameters
                    FileStatus fileStatus = new FileStatus(
                        sUrl,
                        username,
                        domain,
                        GetRootWeb,
                        WebNames,
                        IsSPOnline,
                        UndoCheckOut,
                        DeleteIfNoPrevVersions,
                        writer,
                        oldOut);

                    //Generate the Reports for the Site Collection
                    fileStatus.ProcessSiteCollection();
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



    }
}
