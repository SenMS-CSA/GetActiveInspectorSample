using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace GetActiveInspectorSample
{
    internal class MailItemWrapper : InspectorWrapper
    {

        // The Object instance behind the Inspector, which is the current item.
        public Outlook.MailItem Item { get; private set; }
        private Word.Document wordDoc = null;
        private Word.Application wordApp = null;

        private Outlook.Inspector inspector = null;

        // .ctor
        // <param name="inspector">The Outlook Inspector instance that should be handled.</param>
        public MailItemWrapper(Outlook.Inspector inspector)
            : base(inspector)
        {
            this.inspector = inspector;
        }

        // Method is called when the wrapper has been initialized.
        protected override void Initialize()
        {
            // Get the item in the current Inspector.
            Item = (Outlook.MailItem)Inspector.CurrentItem;

            // Register Item events.
            Item.Open += new Outlook.ItemEvents_10_OpenEventHandler(Item_Open);
            Item.Write += new Outlook.ItemEvents_10_WriteEventHandler(Item_Write);
        }

        // This method is called when the item is visible and the UI is initialized.
        // <param name="Cancel">When you set this property to true, the Inspector is closed.</param>
        void Item_Open(ref bool Cancel)
        {
            // TODO: Implement something 
            wordDoc = (Word.Document)Item.GetInspector.WordEditor;
            wordApp = wordDoc.Application;
            wordApp.WindowSelectionChange -= WordApp_WindowSelectionChange;
            wordApp.WindowSelectionChange += WordApp_WindowSelectionChange;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Sel"></param>
        private void WordApp_WindowSelectionChange(Word.Selection Sel)
        {
            //Note: Want to access properties of the Email (MailItem object) document for which selection change event occured.
            //This event gets trigger for all the opened emails when selection chagne happenes in one particular email. 

            //To access mailItem object for the wordDoc where selection change happened, ActiveInspector() method helps like below. 
            var activeInspector = Globals.ThisAddIn.Application.ActiveInspector();//This doesn't provide correct object inside ribbon callbacks but works here.
            var mailItem = (Outlook.MailItem)inspector.CurrentItem;
            if (activeInspector != inspector)
            {
                System.Windows.Forms.MessageBox.Show("Event NOT for active inspector window. || Selected Text :" + Sel.Text + " || Email Subject: " + mailItem.Subject);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Event for active inspector window. || Selected Text :" + Sel.Text + " || Email Subject: " + mailItem.Subject);
            }
        }

            // This method is called when the item is saved.
            // <param name="Cancel">When set to true, the save operation is cancelled.</param>
        void Item_Write(ref bool Cancel)
        {
            //TODO: Implement something 
        }

        // The Close method is called when the inspector has been closed.
        // Do your cleanup tasks here.
        // The UI is gone, cannot access it here.
        protected override void Close()
        {
            // Unregister events.
            Item.Write -= new Outlook.ItemEvents_10_WriteEventHandler(Item_Write);
            Item.Open -= new Outlook.ItemEvents_10_OpenEventHandler(Item_Open);

            // Release references to COM objects.
            Release(Item);
            Item = null;

            try
            {
                if (wordDoc != null)
                {
                    Marshal.ReleaseComObject(wordDoc);
                }
            }
            catch { }
            wordDoc = null;

            try
            {
                if (wordApp != null)
                {
                    wordApp.WindowSelectionChange -= WordApp_WindowSelectionChange;
                    Marshal.ReleaseComObject(wordApp);
                }
            }
            catch { }

            // Set item to null to keep a reference in memory of the garbage collector.
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void Release(object obj)
        {
            Marshal.ReleaseComObject(obj);
        }
    }
}