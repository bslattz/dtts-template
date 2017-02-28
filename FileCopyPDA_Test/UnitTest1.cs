using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using Microsoft.Win32;

namespace FileCopyPDA_Test
{
    [TestClass]
    public class RegistryQueryTest
    {
        [TestMethod]
        public void FindCurrentVersion ()
        {
            var sut = new FileCopyPDA.FileCopyPDA();

            var value = sut.QueryRegistry(Registry.ClassesRoot,
                @"Excel.Application\CurVer");

            Assert.IsInstanceOfType(value, typeof(string));
        }
        [TestMethod]
        public void GetTemplatePath_Pass()
        {
            var sut = new FileCopyPDA.FileCopyPDA();

            var templatePath = sut.getTemplatePath();

            Assert.IsInstanceOfType(templatePath, typeof(string));
        }

        [TestMethod]
        public void Execute_Pass()
        {
            var args = new AddInPostDeploymentActionArgs(
                new Uri("C:\\Users\\Admin\\Dropbox\\Linda\\addin\\Application Files\\WeekEndingTabs_1_0_0_20\\WeekEndingTabs.dll.manifest"), 
                AddInInstallationStatus.InitialInstall,
                "",
                "", "", "", "WeekEndingTabs", "Test Version", ""
                );
            var sut = new FileCopyPDA.FileCopyPDA();
            sut.Execute(args);
        }
    }
}
