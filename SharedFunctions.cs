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
using System.Xml.Linq;

namespace BrandingTool
{
    static class SharedFunctions
    {
        internal static char[] trimChars = new char[] { '/' };


        public static void UploadMasterPage(ClientContext clientContext, XElement element)
        {
            string rootPath = GetAttribute(element, "rootPath");
            string masterFilePath = GetFullPath(rootPath, GetAttribute(element, "masterFilePath"));
            string previewFilePath = GetFullPath(rootPath, GetAttribute(element, "previewFilePath"));
            var folder = "";  // GetAttribute(element, "folder", false).TrimEnd(trimChars);
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

                string masterFileUrl = UrlUtility.Combine(rootFolder.ServerRelativeUrl,folder, fileName);
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

        public static void UploadPageLayout(ClientContext clientContext, XElement element)
        {
            string rootPath = GetAttribute(element, "rootPath");
            string filePath = GetFullPath(rootPath, GetAttribute(element, "filePath", true));
            string folder = "";  // GetAttribute(element, "folder", false).TrimEnd(trimChars);
            string title = GetAttribute(element, "title");
            string description = GetAttribute(element, "description");
            string associatedContentTypeID = GetAttribute(element, "associatedContentTypeID", true);
            Console.WriteLine("{2} - Uploading {0} to {1}", filePath.Substring(filePath.LastIndexOf('\\') + 1), String.Concat("[Master Page Gallery]/", folder).TrimEnd(trimChars), Environment.NewLine);
            clientContext.Web.DeployPageLayout(filePath, title, description, associatedContentTypeID, folder);
        }

        public static void UploadFile(ClientContext clientContext, XElement element)
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
                    folder = "15/"; //al files in the Theme folder have to be put in the uiVersion folder
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

        public static void UploadTheme(ClientContext clientContext, XElement element)
        {
            string rootPath = GetAttribute(element, "rootPath");
            string themeName = GetAttribute(element, "themeName", true);
            string masterPageName = GetAttribute(element, "masterPageName", true);
            string colorFilePath = GetFullPath(rootPath, GetAttribute(element, "colorFilePath"));
            string backgroundImagePath = GetFullPath(rootPath, GetAttribute(element, "backgroundImagePath"));
            string fontFilePath = GetFullPath(rootPath, GetAttribute(element, "fontFilePath"));
            Console.WriteLine("{1} - Uploading theme and creating Composed Look \"{0}\"", themeName, Environment.NewLine);
            clientContext.Web.DeployThemeToWeb(themeName, colorFilePath, fontFilePath, backgroundImagePath, masterPageName);
        }

        public static void CreateThemeByRelativeUrl(ClientContext clientContext, XElement element)
        {
            string webUrl = GetAttribute(element, "webUrl");
            string themeName = GetAttribute(element, "themeName", true);
            string masterPageName = GetAttribute(element, "masterPageName", true);
            string colorFileUrl = GetAttribute(element, "colorFileUrl");
            string backgroundImageUrl = GetAttribute(element, "backgroundImageUrl");
            string fontFileUrl = GetAttribute(element, "fontFileUrl");
            Console.WriteLine("{1} - Creating Composed Look \"{0}\"", themeName, Environment.NewLine);
            Web destinationWeb = clientContext.Web;
            if (!String.IsNullOrEmpty(webUrl))
            {
                destinationWeb = destinationWeb.GetWeb(webUrl);
                clientContext.Load(destinationWeb);
                clientContext.ExecuteQuery();
            }
            destinationWeb.DeployThemeToWeb(themeName, colorFileUrl, backgroundImageUrl, fontFileUrl, masterPageName);
        }

        public static void ApplyTheme(ClientContext clientContext, XElement element)
        {
            string themeName = GetAttribute(element, "themeName", true);
            string subWebUrl = GetAttribute(element, "subWebUrl");
            bool applyToSubWebs = GetAttribute(element, "applyToSubWebs").ToBoolean();
            var targetWeb = clientContext.Web;
            Console.WriteLine();
            ApplyThemeToWeb(clientContext, themeName, subWebUrl, applyToSubWebs, targetWeb);
        }

        public static void ApplyThemeToWeb(ClientContext clientContext, string themeName, string subWebUrl, bool applyToSubWebs, Web targetWeb)
        {
            if (!String.IsNullOrEmpty(subWebUrl))
            {
                targetWeb = clientContext.Site.OpenWeb(subWebUrl);
            }
            clientContext.Load(targetWeb);
            clientContext.ExecuteQuery();
            Console.WriteLine(" - Applying Composed Look \"{0}\" to {1}", themeName, targetWeb.ServerRelativeUrl);
            targetWeb.SetThemeToSubWeb(clientContext.Web, themeName);
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
                Console.Write("SharePoint Site Url: ");
                if (String.IsNullOrEmpty(strSiteUrl))
                    strSiteUrl = Console.ReadLine();
                else
                    Console.WriteLine(strSiteUrl);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
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
            if (filePath.Length < 2 || filePath.Substring(1, 1) == ":") return filePath; //Already a full path
            return Path.Combine(rootPath, filePath);
        }


        #endregion

    }

        public static partial class Constants
        {
            internal const string HTMLMASTERPAGE_CONTENT_TYPE = "0x0101000F1C8B9E0EB4BE489F09807B2C53288F0054AD6EF48B9F7B45A142F8173F171BD10003D357F861E29844953D5CAA1D4D8A3A";
        }
}
