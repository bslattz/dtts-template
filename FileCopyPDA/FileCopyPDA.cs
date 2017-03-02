using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using Microsoft.VisualStudio.Tools.Applications;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Services = RegistryServices.Service;

namespace FileCopyPDA
{
    public class FileCopyPDA : IAddInPostDeploymentAction
    {
        public interface IMessageBox
        {
            DialogResult Show (IWin32Window window, string text, string caption, 
                MessageBoxButtons buttons);
            DialogResult Show (string text, string caption);
        }
        private class MessageBoxClass : IMessageBox
        {
            public DialogResult Show(IWin32Window window, string text, string caption,
                MessageBoxButtons buttons)
            {
                SetForegroundWindow(window.Handle);
                return MessageBox.Show(window, text, caption, buttons);
            }

            public DialogResult Show(string text, string caption)
            {
                return MessageBox.Show(text, caption);
            }
        }
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetForegroundWindow (IntPtr hWnd);

        private readonly IMessageBox _messageBox;

        public FileCopyPDA()
        {
            _messageBox = new MessageBoxClass();
        }
        public FileCopyPDA (IMessageBox messageBox = null)
        {
            _messageBox = messageBox ?? new MessageBoxClass();
        }
        public void Execute(AddInPostDeploymentActionArgs args)
        {
            switch (args.InstallationStatus)
            {
                case AddInInstallationStatus.InitialInstall:
                case AddInInstallationStatus.Update:
                case AddInInstallationStatus.Uninstall:
                    const string dataDirectory =
                        @"Data\DTTS TEMPLATE AUTO.XP.xlt";
                    const string file = @"DTTS TEMPLATE AUTO.XP.xlt";
                    const string appName = "Excel";
                    var officeVersion =
                        Services.QueryRegistry(Registry.ClassesRoot,
                                @"Excel.Application\CurVer")
                            .Replace(".0", "").Split('.').Last();
                    var destPath = getTemplatePath($"{officeVersion}.0", appName);
                    var destFile = System.IO.Path.Combine(destPath, file);
                    var info = new settingsDictionary(appName);
                    switch (args.InstallationStatus)
                    {
                        case AddInInstallationStatus.InitialInstall:
                        case AddInInstallationStatus.Update:
                            var sourcePath = args.AddInPath;
                            var deploymentManifestUri = args.ManifestLocation;


                            var sourceFile = System.IO.Path.Combine(sourcePath,
                                dataDirectory);

                            var title =
                                $"Install DTTS ver {args.Version} for {info[officeVersion]}";
                            var message =
                                $"Add template to:\n  {destFile}?\nUpdate status: {args.InstallationStatus.ToString()}";
                            if (_messageBox.Show(new Form { TopMost = true },
                                    message, title,
                                    MessageBoxButtons.OKCancel) ==
                                DialogResult.Cancel)
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
                    break;
                case AddInInstallationStatus.RunFromCache:
                    break;
                case AddInInstallationStatus.RunFromFolder:
                    break;
                case AddInInstallationStatus.Offline:
                    break;
                case AddInInstallationStatus.Rollback:
                    break;
            }
        }

        public string getTemplatePath(string ver, string appName)
        {
            var app = new Excel.Application { Visible = false };

            var destPathApp = app.TemplatesPath;
            app.Quit();

            var destPath = Services
                .FindKey(Registry.CurrentUser,
                    new List<string> {"SOFTWARE", "Microsoft", "Office", ver, appName, "Options"})
                    .GetValue("PersonalTemplates") as string;

            var message = $"From the App:\n  {destPathApp}" +
                $"\nFrom the Registry:\n  {destPath}";

            _messageBox.Show(message, "Default User Template Path");

            return destPathApp ?? destPath;
        }
        class settingsDictionary
        {
            public struct OfficeSettings
            {
                public string Name;
                public string PersonalTemplatesRegPath;
                public string UserTemplatesRegPath;

                private const string path =
                    @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\{VER}\{APPLICATION}\Options";

                public OfficeSettings (string name, string personal, string user)
                {
                    Name = name;
                    PersonalTemplatesRegPath = personal;
                    UserTemplatesRegPath = user;
                }
            }

            private string _application;

            public settingsDictionary(string application)
            {
                _application = application;
            }

            private string transform (string s, int i) => 
                s.Replace("{VER}", i.ToString()).Replace("{APPLICATION}", _application);
            public string this[string key] => _officeVer[key];
            private readonly Dictionary<string, string> _officeVer = new Dictionary<string, string>
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
}