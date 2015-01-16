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
        static void Main(string[] args)
        {
            try
            {
                SharedFunctions.RunProgram(args);
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0}ERROR OCCURED{0}", Environment.NewLine);
                Console.WriteLine("{0}{1}", ex.Message, Environment.NewLine);
                //Console.WriteLine("{0}{1}", ex.StackTrace, Environment.NewLine);
                SharedFunctions.ExitProgram();
            }
        }
    }
}
