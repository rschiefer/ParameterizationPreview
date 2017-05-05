using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Web.XmlTransform;

namespace Company.ParameterizationPreview
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidParameterizationPreviewPkgString)]
    //[ProvideAutoLoad("{f1536ef8-92ec-443c-9ed7-fdadf150da82}")]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]

    public sealed class ParameterizationPreviewPackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public ParameterizationPreviewPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidParameterizationPreviewCmdSet, (int)PkgCmdIDList.cmdidPreviewParam);
                var menuItem = new OleMenuCommand(MenuItemCallback_PreviewParameterization, menuCommandID);
                menuItem.BeforeQueryStatus += menuItem_BeforeQueryStatus;
                mcs.AddCommand(menuItem);

                // Create the command for the menu item.
                CommandID menuCommandID2 = new CommandID(GuidList.guidParameterizationPreviewCmdSet, (int)PkgCmdIDList.cmdidPreviewToTranformParam);
                var menuPreviewToTransformItem = new OleMenuCommand(MenuItemCallback_CompareParameterizationToTransform, menuCommandID2);
                menuPreviewToTransformItem.BeforeQueryStatus += menuItem_BeforeQueryStatus;
                mcs.AddCommand(menuPreviewToTransformItem);

            }
        }

        void menuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            var myCommand = sender as OleMenuCommand;
            myCommand.Enabled = false;
            myCommand.Visible = false;
            //myCommand.Text = "NEW NAME";


            DTE dte = (DTE)GetService(typeof(SDTE));
            if (dte.SelectedItems.Count > 0)
            {
                var items = dte.SelectedItems as SelectedItems;
                foreach (SelectedItem item in items)
                {
                    if (item.Name.StartsWith("SetParameters.", StringComparison.OrdinalIgnoreCase))
                    {
                        myCommand.Enabled = true;
                        myCommand.Visible = true;
                    }
                }
            }
        }
        #endregion

        [DllImport("user32")]
        private static extern short GetKeyState(int vKey);

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback_PreviewParameterization(object sender, EventArgs e)
        {
            var myCommand = sender as OleMenuCommand;

            DTE dte = (DTE)GetService(typeof(SDTE));
            if (dte.SelectedItems.Count > 0)
            {
                var items = dte.SelectedItems as SelectedItems;
                foreach (SelectedItem item in items)
                {
                    if (item.Name.StartsWith("SetParameters.", StringComparison.OrdinalIgnoreCase))
                    {
                        var configFile = GetProjectFile(item.ProjectItem, "web.config");
                        configFile = GetProjectFile(item.ProjectItem, "app.config") ?? configFile;

                        var result = GenerateParameterizationResult(item, configFile);

                        if (item.DTE.SourceControl.IsItemUnderSCC(configFile))
                            item.DTE.SourceControl.CheckOutItem(configFile);
                        PrettifyXml(configFile);

                        dte.ExecuteCommand("Tools.DiffFiles", string.Format("\"{0}\" \"{1}\"", configFile, result));
                    }
                }
            }         
        }
        private void MenuItemCallback_CompareParameterizationToTransform(object sender, EventArgs e)
        {
            var myCommand = sender as OleMenuCommand;

            DTE dte = (DTE)GetService(typeof(SDTE));
            if (dte.SelectedItems.Count > 0)
            {
                var items = dte.SelectedItems as SelectedItems;
                foreach (SelectedItem item in items)
                {
                    if (item.Name.StartsWith("SetParameters.", StringComparison.OrdinalIgnoreCase))
                    {
                        var configFile = GetProjectFile(item.ProjectItem, "web.config");
                        configFile = GetProjectFile(item.ProjectItem, "app.config") ?? configFile;

                        var parameterizedResult = GenerateParameterizationResult(item, configFile);
                        PrettifyXml(parameterizedResult);
                        var transformedResult = GenerateConfigTransformResult(item, configFile);

                        dte.ExecuteCommand("Tools.DiffFiles", string.Format("\"{0}\" \"{1}\"", transformedResult, parameterizedResult));
                    }
                }
            }
        }

        private string GenerateParameterizationResult(SelectedItem item, string configFile)
        {
            var fullPath = item.ProjectItem.get_FileNames(0);
            var projectDir = fullPath.Substring(0, fullPath.LastIndexOf("\\"));
            var solutionDir = projectDir.Substring(0, projectDir.LastIndexOf("\\"));
            var packagePath = string.Format("{0}\\temp\\ParameterizationPreview\\package\\package.zip", solutionDir);
            var destPath = string.Format("{0}\\temp\\ParameterizationPreview\\dest", solutionDir);
            var sourcePath = string.Format("{0}\\temp\\ParameterizationPreview\\source", solutionDir);


            var paramTempPath = solutionDir + "\\temp\\ParameterizationPreview";
            RunProcess(string.Format("\"del \"{0}\\*.*\" /q /s /f\"", paramTempPath));

            RunProcess(string.Format("\"mkdir \"{0}/package\"\"", paramTempPath));
            RunProcess(string.Format("\"mkdir \"{0}/source\"\"", paramTempPath));
            RunProcess(string.Format("\"mkdir \"{0}/dest\"\"", paramTempPath));

            RunProcess(string.Format("\"copy /Y \"{0}\\*.config\" \"{1}\"\"", projectDir, sourcePath));

            var msdeployExe = "\"C:\\Program Files (x86)\\IIS\\Microsoft Web Deploy V3\\msdeploy.exe\"";

            var parametersFile = GetProjectFile(item.ProjectItem, "parameters.xml");
            if (parametersFile == null)
            {
                throw new FileNotFoundException("Parameters.xml file must be in the root of the project.  Please add the file and retry.");
            }

            var strDeclareCmdText = string.Format("\"{2} -verb:sync -source:dirPath=\"{0}\" -dest:package=\"{1}\" -declareParamFile:\"{4}\"\"", sourcePath, packagePath, msdeployExe, projectDir, parametersFile);
            RunProcess(strDeclareCmdText);


            var strSetCmdText = string.Format("\"{3} -verb:sync -source:package=\"{0}\" -dest:dirPath=\"{1}\" -setParamFile:\"{2}\"\"", packagePath, destPath, fullPath, msdeployExe);

            RunProcess(strSetCmdText);

            var parameterizedConfig = string.Format("{0}\\{1}", destPath, configFile.Remove(0, configFile.LastIndexOf('\\')));
            PrettifyXml(parameterizedConfig);

            return parameterizedConfig;
        }


        private string GenerateConfigTransformResult(SelectedItem item, string configFile)
        {
            var configuration = item.Name.Substring(item.Name.IndexOf('.') + 1);
            configuration = configuration.Remove(configuration.LastIndexOf('.'));

            var configFilename = configFile.Remove(configFile.LastIndexOf('.')).Substring(configFile.LastIndexOf('\\') + 1);
            var configTransformFilename = string.Format("{0}.{1}.config", configFilename, configuration);
            var configTranformFile = item.ProjectItem.Collection.Item(configFilename + ".config").ProjectItems.Item(configTransformFilename).get_FileNames(0);
            
            var transformation = new XmlTransformation(configTranformFile);

            var xmlSource = new XmlDocument();
            xmlSource.Load(configFile);

            transformation.Apply(xmlSource);
            
            var fullPath = item.ProjectItem.get_FileNames(0);
            var projectDir = fullPath.Substring(0, fullPath.LastIndexOf("\\"));
            var solutionDir = projectDir.Substring(0, projectDir.LastIndexOf("\\"));


            var transformedFolder = string.Format("{0}\\temp\\ParameterizationPreview\\ConfigTransform", solutionDir);
            RunProcess(string.Format("\"del \"{0}\\*.*\" /q /s /f\"", transformedFolder));
            RunProcess(string.Format("\"mkdir \"{0}\"\"", transformedFolder));

            var transformedConfig = string.Format("{0}\\transformed.config", transformedFolder);
            var settings = new XmlWriterSettings
            {
                Indent = true,
                NewLineOnAttributes = true
            };
            using (var xmlWriter = XmlWriter.Create(transformedConfig, settings))
            {
                xmlSource.WriteTo(xmlWriter);
            }

            return transformedConfig;
        }
        void PrettifyXml(string filePath)
        {
            var xml = new XmlDocument();
            xml.Load(filePath);
            var settings = new XmlWriterSettings
            {
                Indent = true,
                NewLineOnAttributes = true
            };
            using (var xmlWriter = XmlWriter.Create(filePath, settings))
            {
                xml.WriteTo(xmlWriter);
            }
        }

        string GetProjectFile(ProjectItem project, string name)
        {
            try
            {
                var file = project.Collection.Item(name);
                return file.get_FileNames(0);
            }
            catch (Exception)
            {
                return null;
            }
        }

        void RunProcess(string command)
        {
            try
            {
                // Check if control key is down
                var controlKeyState = GetKeyState((int)Keys.ControlKey);
                var controlKeyStateBits = BitConverter.GetBytes(controlKeyState);
                var debugMode = controlKeyStateBits[1] > 0;

                // Change /C to /K to leave cmd open for debugging
                var cmdFlag = debugMode ? "/K" : "/C";
                var process = System.Diagnostics.Process.Start("cmd.exe", cmdFlag + " " + command);
                process.WaitForExit(120000);
                process.WriteToDebug();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

    }

    public static class Extensions
    {
        public static System.Diagnostics.Process WriteToDebug(this System.Diagnostics.Process process)
        {

            try
            {
                Debug.Write(process.StandardOutput.ReadToEnd());
            }
            catch (Exception) { }
            return process;
        }
    }
}
