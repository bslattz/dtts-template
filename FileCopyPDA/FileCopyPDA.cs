using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using Microsoft.VisualStudio.Tools.Applications;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace FileCopyPDA
{
    public class FileCopyPDA : IAddInPostDeploymentAction
    {
        public void Execute(AddInPostDeploymentActionArgs args)
        {
            const string dataDirectory = @"Data\DTTS TEMPLATE AUTO.XP.xlt";
            const string file = @"DTTS TEMPLATE AUTO.XP.xlt";
            var destPath = getTemplatePath();
            var destFile = System.IO.Path.Combine(destPath, file);

            switch (args.InstallationStatus)
            {
                case AddInInstallationStatus.InitialInstall:
                case AddInInstallationStatus.Update:
                    var sourcePath = args.AddInPath;
                    var deploymentManifestUri = args.ManifestLocation;

                    var officeVersion = QueryRegistry(Registry.ClassesRoot,
                        @"Excel.Application\CurVer")
                        .Replace(".0", "").Split('.').Last();

                    var sourceFile = System.IO.Path.Combine(sourcePath, dataDirectory);

                    if (MessageBox.Show(new Form { TopMost = true }, 
                        $"Add template to {destFile}?\nUpdate status: {args.InstallationStatus.ToString()}" +
                        $"\nRegistry Templates Path {Environment.GetFolderPath(Environment.SpecialFolder.Templates)}",

                        $"Install DTTS ver {args.Version} for {officeVer[officeVersion]}",
                        MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                        return;

                    File.Copy(sourceFile, destFile, true);
                    ServerDocument.RemoveCustomization(destFile);
                    ServerDocument.AddCustomization(destFile,
                        deploymentManifestUri);
                    break;
                case AddInInstallationStatus.Uninstall:
                    if (File.Exists(destFile))
                    {
                        File.Delete(destFile);
                    }
                    break;
            }
        }

        public string QueryRegistry(RegistryKey root, string path)
        {
            var keys = path.Split(Path.DirectorySeparatorChar);
            return keys.Aggregate(root, (r, k)  => 
                r?.OpenSubKey(k)
            ).GetValue(null).ToString();
        }

        public string getTemplatePath()
        {
            var app = new Excel.Application { Visible = false };

            var destPath = app.TemplatesPath;
            if (string.IsNullOrEmpty(destPath))
                destPath =
                    Environment.GetFolderPath(Environment.SpecialFolder.Templates);
            app.Quit();
            return destPath;
        }
        private Dictionary<string, string> officeVer = new Dictionary<string, string>
        {
            { "7", "Office 97" },
            { "8", "Office 98" },
            { "9", "Office 2000" },
            { "10", "Office XP" },
            { "11", "Office 2003" },
            { "12", "Office 2007" },
            { "14", "Office 2010" },
            { "15", "Office 2013" },
            { "16", "Office 2016" }
        };
    }
}