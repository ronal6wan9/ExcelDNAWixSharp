using System;
using System.Xml.Linq;
using System.Linq;
using WixSharp;
using WixSharp.CommonTasks;
using WixSharp.Forms;
using System.Windows.Forms;
using WixSharp.UI.Forms;

namespace WixSharpSetup
{
    static class Program
    {
        static void Main()
        {
            var info = XDocument.Load(@"wix\SetupInfo.wxi");
            var productName = info.Text("PRODUCT_NAME");
            var companyName = info.Text("COMPANY_NAME");
            var xllDirPath = info.Text("XLL_DIR_PATH");
            var xll32 = info.Text("XLL32");
            var xll64 = info.Text("XLL64");
            var caOpenPath = info.Text("CA_OPEN_PATH");

            var project = new ManagedProject(productName,
                new Dir($@"%ProgramFiles%\{companyName}\{productName}",
                    new File($@"{xllDirPath}\{xll32}"),
                    //new File($@"{xllDirPath}\{xll32}.config"),
                    new File($@"{xllDirPath}\{xll64}"),
                    //new File($@"{xllDirPath}\{xll64}.config"),
                    //new File($@"{xllDirPath}\License.lic"),
                    new File($@"{info.Text("CA_USER_FILE")}"),
                    new File($@"{info.Text("CA_MACH_FILE")}"),
                    new File($@"{caOpenPath}\{info.Text("CA_OPEN_FILE")}").Permanent(),
                    new File($@"{caOpenPath}\README.txt")),
                new Property("AddinFolder", "")
            );
            project.AddXmlInclude(@"wix\SetupInfo.wxi");
            project.AddXmlInclude(@"wix\SetupScope.wxi");
            project.BackgroundImage = @"res\companySetupDialog.bmp";
            project.BannerImage = @"res\companySetupBanner.bmp";
            project.ControlPanelInfo.Comments = "$(var.PRODUCT_DESC)";
            project.ControlPanelInfo.Readme = "$(var.PRODUCT_SITE)/manual";
            project.ControlPanelInfo.HelpLink = "$(var.PRODUCT_SITE)/support";
            project.ControlPanelInfo.UrlInfoAbout = "$(var.PRODUCT_SITE)/about";
            project.ControlPanelInfo.UrlUpdateInfo = "$(var.PRODUCT_SITE)/update";
            project.ControlPanelInfo.ProductIcon = @"res\company.ico";
            project.ControlPanelInfo.Contact = companyName;
            project.ControlPanelInfo.Manufacturer = companyName;
            project.ControlPanelInfo.InstallLocation = "[INSTALLDIR]";
            project.ControlPanelInfo.NoModify = true;

            //The combination of GUID and version will be seed for consistent ProductUpgradeCode and unique ProductId
            project.GUID = new Guid(info.Text("PRODUCT_GUID"));
            project.Include(WixExtension.Util);
            project.Language = info.Text("SETUP_LANG");
            project.LicenceFile = @"res\product_License.rtf";

            //custom set of standard UI dialogs
            project.ManagedUI = new ManagedUI();
            project.ManagedUI.Icon = @"res\company.ico";
            project.ManagedUI.InstallDialogs
                .Add(Dialogs.Welcome)
                .Add(Dialogs.Licence)
                .Add(Dialogs.InstallScope)
                .Add(Dialogs.Progress)
                .Add(Dialogs.Exit);
            project.ManagedUI.ModifyDialogs
                .Add(Dialogs.MaintenanceType)
                .Add(Dialogs.Progress)
                .Add(Dialogs.Exit);

            project.MajorUpgradeStrategy = MajorUpgradeStrategy.Default;
            project.OutDir = @"..\Build\";
            project.SetVersionFromFile(info.Text("VER_DLL_FILE"));
            project.OutFileName = $"{productName}{project.Version}";
#if DEBUG
            project.PreserveTempFiles = true;
#endif
            project.SetNetFxPrerequisite("NETFRAMEWORK40FULL='#1'", $"{productName} requires .NET Framework 4.0.");
            project.UIInitialized += Project_UIInitialized;
            project.ValidateBackgroundImage = false;
            project.BuildMsi();
        }

        private static void Project_UIInitialized(SetupEventArgs e)
        {
            e.ManagedUI.OnCurrentDialogChanged += ManagedUI_OnCurrentDialogChanged;
        }

        private static void ManagedUI_OnCurrentDialogChanged(IManagedDialog obj)
        {
            if (obj.GetType() == Dialogs.InstallScope)
            {
                //Somehow the layout if InstallScope is really different from the rest
                //This mess is an attempt to fix the mess temporarily
                obj._Control("tableLayoutPanel1").Size = new System.Drawing.Size(491, 43);
                obj._Control("back").SetBounds(222, 10, 77, 23);
                obj._Control("next").SetBounds(305, 10, 77, 23);
                obj._Control("cancel").SetBounds(402, 86, 77, 23);

                var upShift = (int)(obj._Control("next").Height * 2.3) - obj._Control("bottomPanel").Height;
                obj._Control("bottomPanel").Top -= upShift;
                obj._Control("bottomPanel").Height += upShift;

                obj._Control("middlePanel").Top = obj._Control("topBorder").Bottom + 5;
                obj._Control("middlePanel").Height = (obj._Control("bottomPanel").Top - 5) - obj._Control("middlePanel").Top;
            }
            else if (obj.GetType() == Dialogs.MaintenanceType)
                obj._Control("change").Enabled = false;
        }

        static string Text(this XDocument inc, string key) =>
            inc.Root.Nodes().OfType<XProcessingInstruction>().First(d => d.Data.Split('=')[0] == key).Data.Split('=')[1].Trim().Trim('"', '\\');

        static File Permanent(this File file)
        {
            file.Attributes = new System.Collections.Generic.Dictionary<string, string> { { "Component:Permanent", "yes" } };
            return file;
        }

        static Property Secure(this Property prop)
        {
            prop.Attributes = new System.Collections.Generic.Dictionary<string, string> { { "Secure", "yes" } };
            return prop;
        }

        static Control _Control(this IManagedDialog obj, string name) =>
            (obj as ManagedForm).Controls.Find(name, true)[0];
    }
}