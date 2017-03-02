using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using Microsoft.Win32;
using IWin32Window = System.Windows.Forms.IWin32Window;
using Services = RegistryServices.Service;

namespace FileCopyPDA_Test
{
    [TestClass]
    public class RegistryQueryTest
    {
        [TestMethod]
        public void GetTemplatePath_Pass()
        {
            const string appName = "Excel";
            var officeVersion = Services.QueryRegistry(Registry.ClassesRoot,
                @"Excel.Application\CurVer")
                .Replace(".0", "").Split('.').Last();
            var mb = new mockMessageBox();
            var sut = new FileCopyPDA.FileCopyPDA();

            var templatePath = sut.getTemplatePath($"{officeVersion}.0", appName);

            Assert.IsInstanceOfType(templatePath, typeof(string));
        }
        [TestMethod]
        public void Execute_Pass()
        {
            var args = new test_AddInPostDeploymentActionArgs(
                new Uri("C:\\Users\\Admin\\Dropbox\\Linda\\addin\\Application Files\\WeekEndingTabs_1_0_0_20\\WeekEndingTabs.dll.manifest"), 
                AddInInstallationStatus.InitialInstall,
                "",
                "", "", "", "WeekEndingTabs", "Test Version", ""
                );
            var mb = new mockMessageBox();
            var sut = new FileCopyPDA.FileCopyPDA(mb);
            sut.Execute(args);
            Debug.WriteLine(string.Join(("\n"),
                mb.Trace.Select(o => $"{o.caption}\t{o.text}\t{o.Result}"))
            );
            Assert.IsTrue(mb.Trace.Any(m => m.caption.StartsWith("Install DTTS ver")));
        }
        public class mockMessageBox : FileCopyPDA.FileCopyPDA.IMessageBox
        {
            public struct Message
            {
                public string text;
                public string caption;
                public DialogResult Result;

                public Message Add(string t, string c, DialogResult r)
                {
                    text = t;
                    caption = c;
                    Result = r;
                    return this;
                }
            }
            public List<Message> Trace = new List<Message>();
            public DialogResult Show (IWin32Window window, string text, string caption,
                MessageBoxButtons buttons)
            {
                var res = MessageBox.Show(window, text, caption, buttons);
                Trace.Add(new Message().Add(text, caption, res));
                FileCopyPDA.FileCopyPDA.SetForegroundWindow(window.Handle);
                return res;               
            }
            public DialogResult Show (string text, string caption)
            {
                var res = MessageBox.Show(text, caption);
                Trace.Add(new Message().Add(text, caption, res));
                return res;
            }
        }
        public class test_AddInPostDeploymentActionArgs :
            AddInPostDeploymentActionArgs
        {
            private List<string> trace { get; } = new List<string>();
            public test_AddInPostDeploymentActionArgs(
                Uri manifestLocation, 
                AddInInstallationStatus installationStatus, 
                string deploymentManifestXml, 
                string applicationManifestXml, 
                string hostManifestXml, 
                string postActionManifestXml, 
                string productName, 
                string version, 
                string addInPath) : base(manifestLocation, installationStatus, deploymentManifestXml, applicationManifestXml, hostManifestXml, postActionManifestXml, productName, version, addInPath)
            {}
        }
    }
}
