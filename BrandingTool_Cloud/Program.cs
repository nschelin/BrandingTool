using System;

namespace BrandingTool
{
    class Program
    {
        static void Main(string[] args)
        {
            //Uncoment the following line for "Release" debugging
            //SharedFunctions.RunProgram(args); return;

            try
            {
                SharedFunctions.RunProgram(args);
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0}ERROR OCCURED{0}", Environment.NewLine);
                Console.WriteLine("{0}{1}", ex.Message, Environment.NewLine);
                SharedFunctions.ExitProgram();
            }
        }
    }
}
