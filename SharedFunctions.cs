using System;
using System.IO;
using System.Security;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;

namespace BrandingTool
{
    static class SharedFunctions
    {
        internal static char[] trimChars = new char[] { '/' };
        internal static string defaultFile = System.AppDomain.CurrentDomain.FriendlyName.Replace(".vshost", "").Replace(".exe", ".xml");


        public static void RunProgram(string[] args, bool OnPrem = false)
        {
            if (OnPrem)
                Console.WriteLine("BRANDING TOOL FOR SHAREPOINT ON PREMISES");
            else
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

                if (OnPrem)
                    Process_OnPrem(defaultRootPath, site, siteUrl, siteUsername, sitePassword);
                else
                    Process_Cloud(defaultRootPath, site, siteUrl, siteUsername, sitePassword);
            }


            Console.WriteLine("{0}{0}Done!", Environment.NewLine);
            SharedFunctions.ExitProgram();
        }

        private static void Process_Cloud(string defaultRootPath, System.Xml.Linq.XElement site, string siteUrl, string siteUsername, string sitePassword)
        {
            Console.WriteLine("Updating Branding at {0}", siteUrl);
            try
            {
                var am = new AuthenticationManager();
                ClientContext cc = am.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, siteUsername, sitePassword);
                cc.Load(cc.Web);
                cc.ExecuteQuery();
                cc.Dispose();
            }
            catch (Exception)
            {
                Console.WriteLine("{1}Login failed for \"{0}\"", siteUsername, Environment.NewLine);
                SharedFunctions.ExitProgram();
            }
            var authManager = new AuthenticationManager();
            using (ClientContext clientContext = authManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, siteUsername, sitePassword))
            {
                clientContext.Load(clientContext.Web);
                clientContext.ExecuteQuery();

                ProcessItems(defaultRootPath, site, clientContext);
            }
        }

        private static void Process_OnPrem(string defaultRootPath, System.Xml.Linq.XElement site, string siteUrl, string siteUsername, string sitePassword)
        {
            Console.WriteLine("Updating Branding at {0}", siteUrl);
            try
            {
                var am = new AuthenticationManager();
                ClientContext cc = am.GetNetworkCredentialAuthenticatedContext(siteUrl, siteUsername.Substring(0, siteUsername.IndexOf("@")), sitePassword, siteUsername.Substring(siteUsername.IndexOf("@") + 1));
                cc.Load(cc.Web);
                cc.ExecuteQuery();
                cc.Dispose();
            }
            catch (Exception)
            {
                Console.WriteLine("{1}Login failed for \"{0}\"", siteUsername, Environment.NewLine);
                SharedFunctions.ExitProgram();
            }
            var authManager = new AuthenticationManager();
            using (ClientContext clientContext = authManager.GetNetworkCredentialAuthenticatedContext(siteUrl, siteUsername.Substring(0, siteUsername.IndexOf("@")), sitePassword, siteUsername.Substring(siteUsername.IndexOf("@") + 1)))
            {
                clientContext.Load(clientContext.Web);
                clientContext.ExecuteQuery();

                ProcessItems(defaultRootPath, site, clientContext);
            }
        }

        private static void ProcessItems(string defaultRootPath, System.Xml.Linq.XElement site, ClientContext clientContext)
        {
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
                        SharedFunctions.CreateTheme(clientContext, element);
                        break;
                    case "applytheme":
                        SharedFunctions.ApplyTheme(clientContext, element);
                        break;
                    default:
                        break;
                }
            }
        }

        private static void UploadMasterPage(ClientContext clientContext, XElement element)
        {
            string rootPath = GetAttribute(element, "rootPath");
            string masterFilePath = GetFullPath(rootPath, GetAttribute(element, "masterFilePath"));
            string previewFilePath = GetFullPath(rootPath, GetAttribute(element, "previewFilePath"));
            var folder = GetAttribute(element, "folder", false).TrimEnd(trimChars);
            if (folder.Length > 0)
                folder += "/";
            var title = GetAttribute(element, "title", true);
            var description = GetAttribute(element, "description");
            var uiVersion = GetAttribute(element, "uiVersion");
            uiVersion = uiVersion == "" ? "15" : uiVersion;
            var defaultCssFile = GetAttribute(element, "defaultCssFile");
            Console.WriteLine("{2} - Uploading {0} to {1}", masterFilePath.Substring(masterFilePath.LastIndexOf('\\') + 1), String.Concat("[Master Page Gallery]/", folder).TrimEnd(trimChars), Environment.NewLine);
            if (!String.IsNullOrEmpty(masterFilePath))
            {
                if (Path.GetExtension(masterFilePath) == ".master")
                {
                    //If there is an .html Master Page present, error out
                    try
                    {
                        Microsoft.SharePoint.Client.File file2Delete = clientContext.Web.GetFileByServerRelativeUrl(String.Concat(clientContext.Web.ServerRelativeUrl, "/_catalogs/masterpage/", folder, Path.GetFileNameWithoutExtension(masterFilePath), ".html"));
                        clientContext.Load(file2Delete);
                        clientContext.ExecuteQuery();
                        if (file2Delete.Exists)
                        {
                            Console.WriteLine(String.Concat("{0}\tERROR: The Master Page \"{1}\" {0}",
                                                            "\thas an associated .html file and canot be updated. Use {0}",
                                                            "\tan HTML Master Page (.html) to recreate the .master file {0}",
                                                            "\tor delete \"{2}\"{0}")
                                                            , Environment.NewLine, Path.Combine(folder, Path.GetFileName(masterFilePath)), Path.Combine(folder, Path.GetFileName(file2Delete.Name)));
                            ExitProgram();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
                clientContext.Web.DeployMasterPage(masterFilePath, title, description, uiVersion, defaultCssFile, folder);
                //Change Content Type to HTML Master Page
                var web = clientContext.Web;
                string fileName = Path.GetFileName(masterFilePath);

                // Get the path to the file which we are about to deploy
                List masterPageGallery = web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                Folder rootFolder = masterPageGallery.RootFolder;
                web.Context.Load(masterPageGallery);
                web.Context.Load(rootFolder);
                web.Context.ExecuteQuery();

                string masterFileUrl = UrlUtility.Combine(rootFolder.ServerRelativeUrl, folder, fileName);
                Microsoft.SharePoint.Client.File masterFile = web.GetFileByServerRelativeUrl(masterFileUrl);
                web.Context.Load(masterFile);
                web.Context.ExecuteQuery();

                var listItem = masterFile.ListItemAllFields;
                if (masterPageGallery.ForceCheckout || masterPageGallery.EnableVersioning)
                {
                    if (masterFile.CheckOutType == CheckOutType.None)
                    {
                        masterFile.CheckOut();
                    }
                }

                // Set content type as master page
                listItem["ContentTypeId"] = Constants.HTMLMASTERPAGE_CONTENT_TYPE;
                listItem.Update();
                if (masterPageGallery.ForceCheckout || masterPageGallery.EnableVersioning)
                {
                    masterFile.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                    listItem.File.Publish(string.Empty);
                }
                web.Context.Load(listItem);
                web.Context.ExecuteQuery();
            }
            if (!String.IsNullOrEmpty(previewFilePath))
            {
                Console.WriteLine("{2}   Uploading {0} to {1}", previewFilePath.Substring(previewFilePath.LastIndexOf('\\') + 1), String.Concat("[Master Page Gallery]/", folder).TrimEnd(trimChars), String.IsNullOrEmpty(masterFilePath) ? Environment.NewLine : "");
                List library = clientContext.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                clientContext.Load(library);
                clientContext.ExecuteQuery();
                var destFolder = library.RootFolder;
                if (folder.Length > 0)
                {
                    destFolder = clientContext.Web.EnsureFolder(library.RootFolder, folder);
                }
                destFolder.UploadFile(previewFilePath);
            }
        }

        private static void UploadPageLayout(ClientContext clientContext, XElement element)
        {
            string rootPath = GetAttribute(element, "rootPath");
            string filePath = GetFullPath(rootPath, GetAttribute(element, "filePath", true));
            string folder = GetAttribute(element, "folder", false).TrimEnd(trimChars);
            string title = GetAttribute(element, "title");
            string description = GetAttribute(element, "description");
            string associatedContentTypeID = GetAttribute(element, "associatedContentTypeID", true);
            Console.WriteLine("{2} - Uploading {0} to {1}", filePath.Substring(filePath.LastIndexOf('\\') + 1), String.Concat("[Master Page Gallery]/", folder).TrimEnd(trimChars), Environment.NewLine);
            clientContext.Web.DeployPageLayout(filePath, title, description, associatedContentTypeID, folder);
        }

        private static void UploadFile(ClientContext clientContext, XElement element)
        {
            string rootPath = GetAttribute(element, "rootPath");
            string filePath = GetFullPath(rootPath, GetAttribute(element, "filePath", true));
            var folder = GetAttribute(element, "folder", true).TrimEnd(trimChars);
            var libraryName = GetAttribute(element, "library", true).TrimEnd(trimChars).ToLower();
            List library;
            switch (libraryName)
            {
                case "[masterpage]":
                    libraryName = "[Master Page Gallery]";
                    library = clientContext.Web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
                    clientContext.Load(library);
                    clientContext.ExecuteQuery();
                    break;
                case "[theme]":
                    folder = "15/"; //all files in the Theme folder have to be put in the uiVersion folder
                    libraryName = "[Theme Gallery]";
                    library = clientContext.Web.GetCatalog((int)ListTemplateType.ThemeCatalog);
                    clientContext.Load(library);
                    clientContext.ExecuteQuery();
                    break;
                default:
                    library = clientContext.Web.GetList(libraryName);
                    clientContext.Load(library);
                    clientContext.ExecuteQuery();
                    break;
            }
            var fileName = Path.GetFileName(filePath);
            var searchPattern = (String.IsNullOrEmpty(fileName) || fileName == "*") ? "*.*" : fileName;
            var searchOption = (String.IsNullOrEmpty(fileName) || fileName == "*") ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            string[] fileList = Directory.GetFiles(Path.GetDirectoryName(filePath), searchPattern, searchOption);
            string rootFolder = Path.GetDirectoryName(filePath);
            foreach (var file in fileList)
            {
                var destFolder = library.RootFolder;
                string fullFolder = folder;
                if (Path.GetDirectoryName(file) == rootFolder)
                {
                    //if filename is "*" don't upload the files in the root, only subfolders
                    if (fileName == "*") continue;
                }
                else
                {
                    string newFolders = file.Substring(rootFolder.Length + 1, file.Length - rootFolder.Length - Path.GetFileName(file).Length - 1).Replace("\\", "/").TrimEnd(trimChars);
                    fullFolder = String.Concat(String.IsNullOrEmpty(folder) ? "" : folder + "/", newFolders);
                    if (fullFolder.Length > 0)
                    {
                        destFolder = clientContext.Web.EnsureFolder(library.RootFolder, fullFolder);
                    }
                }
                Console.WriteLine("{2} - Uploading {0} to {1}", Path.GetFileName(file), String.Concat(libraryName, "/", fullFolder).TrimEnd(trimChars), Environment.NewLine);
                destFolder.UploadFile(file);
            }
        }

        private static void UploadTheme(ClientContext clientContext, XElement element)
        {
            string rootPath = GetAttribute(element, "rootPath");
            string themeName = GetAttribute(element, "themeName", true);
            string masterPageName = GetAttribute(element, "masterPageName", true);
            string colorFilePath = GetFullPath(rootPath, GetAttribute(element, "colorFilePath"));
            string backgroundImagePath = GetFullPath(rootPath, GetAttribute(element, "backgroundImagePath"));
            string fontFilePath = GetFullPath(rootPath, GetAttribute(element, "fontFilePath"));
            Console.WriteLine("{1} - Uploading theme files and creating Composed Look \"{0}\"", themeName, Environment.NewLine);
            if (colorFilePath != "")
            {
                clientContext.Web.UploadThemeFile(colorFilePath);
                colorFilePath = UrlUtility.Combine(clientContext.Web.ServerRelativeUrl, string.Format(Constants.THEMES_DIRECTORY, Path.GetFileName(colorFilePath)));
            }
            if (fontFilePath != "")
            {
                clientContext.Web.UploadThemeFile(fontFilePath);
                fontFilePath = UrlUtility.Combine(clientContext.Web.ServerRelativeUrl, string.Format(Constants.THEMES_DIRECTORY, Path.GetFileName(fontFilePath)));
            }
            if (backgroundImagePath != "")
            {
                clientContext.Web.UploadThemeFile(backgroundImagePath);
                backgroundImagePath = UrlUtility.Combine(clientContext.Web.ServerRelativeUrl, string.Format(Constants.THEMES_DIRECTORY, Path.GetFileName(backgroundImagePath)));
            }
            masterPageName = UrlUtility.Combine(clientContext.Web.ServerRelativeUrl, string.Format(Constants.MASTERPAGE_DIRECTORY, masterPageName));
            CreateComposedLook(clientContext, themeName, masterPageName, colorFilePath, backgroundImagePath, fontFilePath);
        }

        private static void CreateTheme(ClientContext clientContext, XElement element)
        {
            string themeName = GetAttribute(element, "themeName", true);
            string masterPageName = GetAttribute(element, "masterPageName", true);
            string colorFilePath = GetAttribute(element, "colorFileUrl");
            string backgroundImagePath = GetAttribute(element, "backgroundImageUrl");
            string fontFilePath = GetAttribute(element, "fontFileUrl");
            Console.WriteLine("{1} - Creating Composed Look \"{0}\"", themeName, Environment.NewLine);
            if (colorFilePath != "")
            {
                colorFilePath = UrlUtility.Combine(clientContext.Web.ServerRelativeUrl, string.Format(Constants.THEMES_DIRECTORY, Path.GetFileName(colorFilePath)));
            }
            if (fontFilePath != "")
            {
                fontFilePath = UrlUtility.Combine(clientContext.Web.ServerRelativeUrl, string.Format(Constants.THEMES_DIRECTORY, Path.GetFileName(fontFilePath)));
            }
            if (backgroundImagePath != "")
            {
                backgroundImagePath = UrlUtility.Combine(clientContext.Web.ServerRelativeUrl, backgroundImagePath);
            }
            masterPageName = UrlUtility.Combine(clientContext.Web.ServerRelativeUrl, string.Format(Constants.MASTERPAGE_DIRECTORY, masterPageName));
            CreateComposedLook(clientContext, themeName, masterPageName, colorFilePath, backgroundImagePath, fontFilePath);
        }

        private static void CreateComposedLook(ClientContext clientContext, string themeName, string masterPageName, string colorFilePath, string backgroundImagePath, string fontFilePath)
        {
            clientContext.Web.CreateComposedLookByUrl(themeName, colorFilePath, fontFilePath, backgroundImagePath, masterPageName);
        }

        private static void ApplyTheme(ClientContext clientContext, XElement element)
        {
            string themeName = GetAttribute(element, "themeName", true);
            string subWebUrl = GetAttribute(element, "subWebUrl");
            bool applyToSubWebs = GetAttribute(element, "applyToSubWebs").ToBoolean();
            var targetWeb = clientContext.Web;
            Console.WriteLine();
            ApplyThemeToWeb(clientContext, themeName, subWebUrl, applyToSubWebs, targetWeb);
        }

        private static void ApplyThemeToWeb(ClientContext clientContext, string themeName, string subWebUrl, bool applyToSubWebs, Web targetWeb)
        {
            if (!String.IsNullOrEmpty(subWebUrl))
            {
                targetWeb = clientContext.Site.OpenWeb(subWebUrl);
            }
            var rootWeb = clientContext.Site.RootWeb;
            clientContext.Load(targetWeb);
            clientContext.Load(rootWeb);
            clientContext.ExecuteQuery();
            if (!(rootWeb.Url == targetWeb.Url || targetWeb.WebTemplate != "CMSPUBLISHING")) return;
            Console.WriteLine(" - Applying Composed Look \"{0}\" to {1}", themeName, targetWeb.ServerRelativeUrl);
            if (!targetWeb.ComposedLookExists(themeName))
            {
                if (rootWeb.ComposedLookExists(themeName))
                {
                    var cLook = clientContext.Site.RootWeb.GetComposedLook(themeName);
                    targetWeb.CreateComposedLookByUrl(themeName, cLook.Theme, cLook.Font, cLook.BackgroundImage, cLook.MasterPage);
                }
                else
                {
                    Console.WriteLine("Composed Look \"{0}\" not found.", themeName);
                }
            }
            targetWeb.SetComposedLookByUrl(themeName);
            if (applyToSubWebs)
            {
                WebCollection webs = targetWeb.Webs;
                clientContext.Load(webs);
                clientContext.ExecuteQuery();
                foreach (var web in webs)
                {
                    ApplyThemeToWeb(clientContext, themeName, "", applyToSubWebs, web);
                }
            }
        }

        #region "helper functions"
        public static void ExitProgram()
        {
            Console.Write("{0}{0}Press [Enter] to exit program . . . ", Environment.NewLine);
            Console.ReadLine();
            Environment.Exit(0);
        }

        public static string GetSiteUrl(string strSiteUrl = "")
        {
            try
            {
                if (String.IsNullOrEmpty(strSiteUrl))
                {
                    Console.Write("SharePoint Site Url: ");
                    strSiteUrl = Console.ReadLine();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("SharePoint Site Url: {0}", e.Message);
                strSiteUrl = string.Empty;
            }
            return strSiteUrl;
        }

        public static string GetUserName(string strUserName = "")
        {
            try
            {
                Console.Write("SharePoint Username: ");
                if (String.IsNullOrEmpty(strUserName))
                    strUserName = Console.ReadLine();
                else
                    Console.WriteLine(strUserName);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                strUserName = string.Empty;
            }
            return strUserName;
        }

        public static SecureString GetPassword(string strPwd = "")
        {
            SecureString sStrPwd = new SecureString();

            try
            {
                Console.Write("SharePoint Password: ");
                if (String.IsNullOrEmpty(strPwd))
                {
                    for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
                    {
                        if (keyInfo.Key == ConsoleKey.Backspace)
                        {
                            if (sStrPwd.Length > 0)
                            {
                                sStrPwd.RemoveAt(sStrPwd.Length - 1);
                                Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                                Console.Write(" ");
                                Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                            }
                        }
                        else if (keyInfo.Key != ConsoleKey.Enter)
                        {
                            Console.Write("*");
                            sStrPwd.AppendChar(keyInfo.KeyChar);
                        }

                    }
                    Console.WriteLine("");
                }
                else
                {
                    Array.ForEach(strPwd.ToCharArray(), sStrPwd.AppendChar);
                    Console.Write(new String('*', (int)(strPwd.Length * 1.8)));
                    Console.WriteLine();
                }
            }
            catch (Exception e)
            {
                sStrPwd = null;
                Console.WriteLine(e.Message);
            }

            return sStrPwd;
        }

        public static string GetAttribute(XElement element, string attribute, bool required = false)
        {
            if (element.Attribute(attribute) == null)
            {
                if (required)
                {
                    Console.WriteLine("ERROR: \"{0}\" is a required element for \"{1}\".", attribute, element.Name);
                    ExitProgram();
                }
                return "";
            }
            return element.Attribute(attribute).Value;
        }

        public static string GetFullPath(string rootPath, string filePath)
        {
            if (filePath.Length < 2 || filePath.Substring(1, 1) == ":" | filePath.StartsWith("\\\\")) return filePath; //Already a full path
            return Path.Combine(rootPath, filePath.TrimStart('\\'));
        }


        #endregion

    }

    public static partial class Constants
    {
        internal const string HTMLMASTERPAGE_CONTENT_TYPE = "0x0101000F1C8B9E0EB4BE489F09807B2C53288F0054AD6EF48B9F7B45A142F8173F171BD10003D357F861E29844953D5CAA1D4D8A3A";
        internal const string THEMES_DIRECTORY = "/_catalogs/theme/15/{0}";
        internal const string MASTERPAGE_DIRECTORY = "/_catalogs/masterpage/{0}";

    }
}
