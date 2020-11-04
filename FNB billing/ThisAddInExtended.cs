using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FNB_billing
{
    public partial class ThisAddIn
    {
        readonly Microsoft.ApplicationInsights.TelemetryClient tc = new Microsoft.ApplicationInsights.TelemetryClient();
        internal void LogTrackInfo(string MenuItem) // Use Azure application insights
        {
            // https://carldesouza.com/how-to-create-custom-events-metrics-traces-in-azure-application-insights-using-c/
            // install the Microsoft.ApplicationInsights NuGet package
            string UserName;
            string PubVer;
            Excel.Application xlAp;
            Excel.Workbook XlWb;
            var EventProperties = new Dictionary<string, string>();
            xlAp = Globals.ThisAddIn.Application;
            XlWb = xlAp.ActiveWorkbook;
            EventProperties.Add("FilePath", XlWb.FullName);
            UserName = System.Environment.GetEnvironmentVariable("username");
            PubVer = "";
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                PubVer = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4); // Returns 4 components i.e. major.minor.build.revision
            }

            tc.InstrumentationKey = "b6d89ab7-9df1-444b-8456-13eebdc85fe7";
            tc.Context.Session.Id = Guid.NewGuid().ToString();
            tc.Context.Device.OperatingSystem = Environment.OSVersion.ToString();
            tc.Context.User.AuthenticatedUserId = UserName;
            tc.Context.Component.Version = PubVer;
            tc.TrackEvent(MenuItem, EventProperties);
            tc.Flush();
        }
        internal void AboutFNB()
        {
            string Msg, PubVer;
            PubVer = "";
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                PubVer = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }

            Msg = "This addin helps with FNB billing.\r\nWritten by Kobus Burger 083 228 9674 ©\r\nVersion: " + PubVer;
            MessageBox.Show(Msg, "FNB Billing");
        }
        internal void ExMsg(Exception Ex)
        {
            Excel.Application xlAp = Globals.ThisAddIn.Application;
            string ErrorDescription = "";
            xlAp.StatusBar = false;
            xlAp.ScreenUpdating = true;
            ErrorDescription = Ex.Source +
                "\r\n0x" + Ex.HResult.ToString("x") + ": " + Ex.Message +
                "\r\n" + Ex.StackTrace +
                "\r\n" + Ex.TargetSite;

            MessageBox.Show(ErrorDescription, "Exception (copy text with Ctrl+C)");
        }


    }

}