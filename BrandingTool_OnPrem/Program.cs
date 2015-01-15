using System;
using System.IO;
using System.Security;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using OfficeDevPnP.Core;
using System.Linq;
using System.Text;
using System.Xml;

namespace BrandingTool
{
    class Program
    {
        internal static char[] trimChars = new char[] { '/' };
        internal static string defaultFile = System.AppDomain.CurrentDomain.FriendlyName.Replace(".vshost","").Replace(".exe", ".xml");

        static void Main(string[] args)
        {

            Console.WriteLine("BRANDING TOOL FOR SHAREPOINT ONLINE(OFFICE 365)");
            Console.WriteLine("   by Don Kirkham{0}", Environment.NewLine);

            string settingsFile = defaultFile;
            if (args.Length > 0)
            {
                settingsFile = args[0];
            }
            if (!System.IO.File.Exists(settingsFile))
            {
                Console.WriteLine("Settings file not found: {0}\r\n", settingsFile);
                Console.WriteLine(String.Concat("\tThe Settings file is a special XML file that can be {0}",
                                                "\tpassed as a command line parameter. The default file is {0}",
                                                "\t\"{1}\" located in the same folder where {0}",
                                                "\t\"{2}\" is executed from.")
                                                , Environment.NewLine, defaultFile, System.AppDomain.CurrentDomain.FriendlyName.Replace(".vshost", ""));
                SharedFunctions.ExitProgram();
            }
            Console.WriteLine("Settings File: {0}{1}", settingsFile, Environment.NewLine);

            var branding = XDocument.Load(settingsFile).Element("branding");
            if (branding == null)
            {
                Console.WriteLine("Settings file not valid: {0}\r\n", settingsFile);
                Console.WriteLine("\tThe settings file must have a \"<branding>\" node.");
                SharedFunctions.ExitProgram();
            }
            string defaultRootPath = Path.GetFullPath(settingsFile);
            defaultRootPath = defaultRootPath.Substring(0, defaultRootPath.LastIndexOf("\\") + 1);

            string defaultUsername = "";
            string defaultPassword = "";
            var defaultCredentials = branding.Element("credentials");
            if (defaultCredentials != null)
            {
                defaultUsername = defaultCredentials.Attribute("username") == null ? "" : defaultCredentials.Attribute("username").Value;
                defaultPassword = defaultCredentials.Attribute("password") == null ? "" : defaultCredentials.Attribute("password").Value;
            }
            foreach (var site in branding.Descendants("site"))
            {
                var siteUrl = SharedFunctions.GetSiteUrl(SharedFunctions.GetAttribute(site, "url")).TrimEnd(trimChars) + "/";
                site.Attribute("url").SetValue(siteUrl);
                var siteUsername = site.Attribute("username") == null ? defaultUsername : site.Attribute("username").Value;
                var sitePassword = site.Attribute("password") == null ? defaultPassword : site.Attribute("password").Value;

                string rootPath = SharedFunctions.GetAttribute(site, "rootPath");
                if (rootPath != "") 
                    defaultRootPath = rootPath;

                Console.WriteLine("Updating Branding at {0}", siteUrl);
                try
                {
                    var am = new AuthenticationManager();
                    var cc = am.GetNetworkCredentialAuthenticatedContext(siteUrl, siteUsername.Substring(0, siteUsername.IndexOf("@")), sitePassword, siteUsername.Substring(siteUsername.IndexOf("@") + 1));
                    cc.ExecuteQuery();
                    cc.Dispose();
                }
                catch (Exception)
                {
                    Console.WriteLine("{1}Login failed for \"{0}\"",siteUsername, Environment.NewLine);
                    SharedFunctions.ExitProgram();
                }

                var authManager = new AuthenticationManager();
                using (ClientContext clientContext = authManager.GetNetworkCredentialAuthenticatedContext(siteUrl, siteUsername.Substring(0, siteUsername.IndexOf("@")), sitePassword, siteUsername.Substring(siteUsername.IndexOf("@")+1)))
                {
                    clientContext.Load(clientContext.Web);
                    clientContext.ExecuteQuery();

                    var tasklist = site;
                    //if (((XElement)site.FirstNode).Name.LocalName.ToLower() == "undobranding")
                    //{
                    //    site.FirstNode.Remove();
                    //    var reverseTasks = new XElement("site");
                    //    //reverseTasks.FirstNode.Attribute("url") = tasklist.Attribute("url");
                    //    Stack<XElement> tasks = new Stack<XElement>(); ;
                    //    foreach (var node in site.Descendants())
                    //    {
                    //        tasks.Push(node);
                    //    }
                    //    foreach (var node in tasks)
                    //    {
                    //        reverseTasks.Add(node);
                    //    }
                    //    tasklist = reverseTasks;
                    //    Console.WriteLine("\r\n\r\n\tHaven't written the undo function yet, but it's coming!");
                    //    //ExitProgram();
                    //}

                    foreach (XElement element in tasklist.Descendants())
                    {
                        element.SetAttributeValue("rootPath", defaultRootPath);
                        switch (element.Name.LocalName.ToLower())
                        {
                            case "uploadmasterpage":
                                SharedFunctions.UploadMasterPage(clientContext, element);
                                break;
                            case "uploadpagelayout":
                                SharedFunctions.UploadPageLayout(clientContext, element);
                                break;
                            case "uploadfile":
                                SharedFunctions.UploadFile(clientContext, element);
                                break;
                            case "uploadtheme":
                                SharedFunctions.UploadTheme(clientContext, element);
                                break;
                            case "createtheme":
                                SharedFunctions.CreateThemeByRelativeUrl(clientContext, element);
                                break;
                            case "applytheme":
                                SharedFunctions.ApplyTheme(clientContext, element);
                                break;
                            default:
                                break;
                        }
                    }
                }
            }


            Console.WriteLine("{0}{0}Done!", Environment.NewLine);
            SharedFunctions.ExitProgram();
        }

    }
}
