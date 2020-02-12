using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GetActiveInspectorSample
{
    internal class AppointmentItemWrapper : InspectorWrapper
    {
        public Outlook.AppointmentItem Item { get; private set; }

        public AppointmentItemWrapper(Outlook.Inspector inspector)
           : base(inspector)
        {
        }

        protected override void Initialize()
        {
            // Get the item in the current Inspector.
            Item = (Outlook.AppointmentItem)Inspector.CurrentItem;

            // Register Item events.
            Item.Open += new Outlook.ItemEvents_10_OpenEventHandler(Item_Open);
            Item.Write += new Outlook.ItemEvents_10_WriteEventHandler(Item_Write);
        }

        void Item_Open(ref bool Cancel)
        {
            // TODO: Implement something 
        }
        void Item_Write(ref bool Cancel)
        {
            //TODO: Implement something 
        }

        protected override void Close()
        {
            // Unregister events.
            Item.Write -= new Outlook.ItemEvents_10_WriteEventHandler(Item_Write);
            Item.Open -= new Outlook.ItemEvents_10_OpenEventHandler(Item_Open);

            // Release references to COM objects.
            Release(Item);
            Item = null;

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
