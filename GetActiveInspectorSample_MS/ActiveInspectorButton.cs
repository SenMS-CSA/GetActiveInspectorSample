using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ActiveInspector();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace GetActiveInspectorSample
{
    [ComVisible(true)]
    public class ActiveInspectorButton : Office.IRibbonExtensibility
    {
        public static string m_CRMBtn_ID;

        public Office.IRibbonUI RibbonUI { get; set; }

        public ActiveInspectorButton()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("GetActiveInspectorSample.ActiveInspectorButton.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            RibbonUI = ribbonUI;
        }

        public string GetIdOfRibbonControl(object ctrl)
        {
            if (ctrl == null) return string.Empty;
            return ((Office.IRibbonControl)ctrl).Id;
        }



        public bool IsEnabled(object control)
        {
            if (m_CRMBtn_ID == null)
            {
                try
                {
                    m_CRMBtn_ID = GetIdOfRibbonControl(control);
                }
                catch (Exception exception)
                {
                }
            }
            return true;
        }

        public bool IsVisible(object control)
        {
            //Commented to try out Microsoft suggested approach below - 12-Jan-2020
            //Outlook.Inspector inspector = GetActiveInspector();

            //Start: Microsoft suggested approach to get inspector using control object.
            Outlook.Inspector inspector = null;
            if (control is Microsoft.Office.Core.IRibbonControl)
            {
                Microsoft.Office.Core.IRibbonControl ribbonControl = control as Microsoft.Office.Core.IRibbonControl;

                if (ribbonControl.Context is Outlook.Inspector)
                {
                    inspector = ribbonControl.Context as Outlook.Inspector;
                }
            }
            //End: Microsoft suggested approach to get inspector using control object.

            if (inspector == null)
                return false;

            if (inspector != null && inspector.CurrentItem != null)
            {
                Outlook.AppointmentItem apt = inspector.CurrentItem as Outlook.AppointmentItem;
                if (apt != null)
                    System.Windows.Forms.MessageBox.Show("Appointment item");

                Outlook.MailItem mail = inspector.CurrentItem as Outlook.MailItem;
                if (mail != null)
                    System.Windows.Forms.MessageBox.Show("Mail item");
            }


            return true;
        }

        public void BtnShowPressed(object control, bool bPressed)
        {
            //var outlook = GetActiveInspector();

            int x = 1;
            //try
            //{
            //    inspectorWrapper.InvalidateRibbonControl(PluginMain.m_CRMBtn_ID);
            //}
            //catch (Exception)
            //{
            //}
        }

        #endregion

        //private Outlook.Inspector GetActiveInspector()
        //{
        //    var outlookApp = Globals.ThisAddIn.Application;
        //    if (outlookApp == null)
        //        return null;

        //    if (outlookApp.ActiveInspector() == null)
        //        return null;
        //    return outlookApp.ActiveInspector();
        //}

        /// <summary>
        /// Solution provided by microsoft
        /// </summary>
        /// <returns></returns>
        private Outlook.Inspector GetActiveInspector()
        {

            //if (Globals.ThisAddIn.Application.ActiveInspector() == null && Globals.ThisAddIn.Application.Inspectors.Count > 0)
            //    return Globals.ThisAddIn.Application.Inspectors[1];
            //else
            //    return Globals.ThisAddIn.Application.ActiveInspector();
            var activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
            var activeInspector = activeWindow as Outlook.Inspector;
            if (activeInspector != null) return activeInspector;
            else
            {
                System.Windows.Forms.MessageBox.Show("got null inspector");
                return null;
            }
        }



        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
