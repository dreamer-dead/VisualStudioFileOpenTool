//Inspired by http://stackoverflow.com/questions/350323/open-a-file-in-visual-studio-at-a-specific-line-number

using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace VisualStudioFileOpenTool
{
	class Program
	{
		[STAThread]
		static void Main(string[] args)
		{
            if (args.Length < 2)
            {
                Console.WriteLine("usage: <version> <file path> <line number>");
                return;
            }

			try
			{
                int vsVersion = 0;
                if (!int.TryParse(args[0], out vsVersion))
                {
                    Console.Error.WriteLine("Failed to pased version!");
                    return;
                }

                string vsString = GetVersionString(vsVersion);
                if (String.IsNullOrEmpty(vsString))
                {
                    Console.Error.WriteLine("Unsupported version os MSVS!");
                    return;
                }

                String filename = args[1];
                EnvDTE80.DTE2 dte2;
                dte2 = (EnvDTE80.DTE2)System.Runtime.InteropServices.Marshal.GetActiveObject(vsString);
                dte2.MainWindow.Activate();
                EnvDTE.Window w = dte2.ItemOperations.OpenFile(filename, EnvDTE.Constants.vsViewKindTextView);

                if (args.Length >= 3)
                {
                    int fileline = 0;
                    if (int.TryParse(args[2], out fileline))
                    {
                        ((EnvDTE.TextSelection)dte2.ActiveDocument.Selection).GotoLine(fileline, true);
                    }
                }
			}
			catch (Exception e)
			{
				Console.Write(e.Message);
			}
		}

		public static string GetHelpMessage()
		{
			var versions = new int [] { 2, 3, 5, 8, 10, 12, 13 };
			string s = "Trying to open specified file at spicified line in active Visual Studio \n\n";

			s += "usage: <version> <file path> <line number> \n\n";

			s += String.Format("{0} {1,21} \n", "Visual Studio version", "value");
			foreach (int version in versions)
			{
				s += String.Format("{0}{1:D2} ", "VisualStudio 20", version);
				s += String.Format("{0,21} \n", version);
			}

			s += "";

			return s;
		}

		public static string GetVersionString(int visualStudioVersionNumber)
		{
			//  Source: http://www.mztools.com/articles/2011/MZ2011011.aspx
			switch (visualStudioVersionNumber)
			{
                case 13:
                    return "VisualStudio.DTE.12.0";
				case 12:
					return "VisualStudio.DTE.11.0";
				case 10:
					return "VisualStudio.DTE.10.0";
				case 8:
					return "VisualStudio.DTE.9.0";
				case 5:
					return "VisualStudio.DTE.8.0";
				case 3:
					return "VisualStudio.DTE.7.1";
				case 2:
					return "VisualStudio.DTE.7";
			}

			return String.Empty;
		}
	}
}
