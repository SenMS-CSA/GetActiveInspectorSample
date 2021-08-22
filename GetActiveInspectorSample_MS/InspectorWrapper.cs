using System;
using Office = Microsoft.Office.Core;
using System.Reflection;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;
using Outlook = Microsoft.Office.Interop.Outlook; 

namespace GetActiveInspectorSample
{
    public delegate void InspectorWrapperClosedEventHandler(Guid id);

    public class InspectorWrapper
    {
        private object _ribbonUi;
       // private Inspector _inspector;

        public event InspectorWrapperClosedEventHandler Closed;

        public Guid Id { get; private set; }

        //public Inspector CurrentInspector
        //{
        //    get
        //    {
        //        return _inspector;
        //    }
        //    set
        //    {
        //        _inspector = value;
        //    }
        //}

        public Office.IRibbonUI RibbonUI { get; set; }

        public Inspector Inspector { get; private set; }

        public InspectorWrapper()
        {
        }

        public InspectorWrapper(Inspector inspector)
        {
            // CurrentInspector = inspector;
            Id = Guid.NewGuid();
            Inspector = inspector;
            // Register Inspector events here
            ((Outlook.InspectorEvents_10_Event)Inspector).Close +=
                new Outlook.InspectorEvents_10_CloseEventHandler(Inspector_Close);
            ((Outlook.InspectorEvents_10_Event)Inspector).Activate +=
                new Outlook.InspectorEvents_10_ActivateEventHandler(Activate);
            ((Outlook.InspectorEvents_10_Event)Inspector).Deactivate +=
                new Outlook.InspectorEvents_10_DeactivateEventHandler(Deactivate);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMaximize +=
                new Outlook.InspectorEvents_10_BeforeMaximizeEventHandler(BeforeMaximize);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMinimize +=
                new Outlook.InspectorEvents_10_BeforeMinimizeEventHandler(BeforeMinimize);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMove +=
                new Outlook.InspectorEvents_10_BeforeMoveEventHandler(BeforeMove);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeSize +=
                new Outlook.InspectorEvents_10_BeforeSizeEventHandler(BeforeSize);
            ((Outlook.InspectorEvents_10_Event)Inspector).PageChange +=
                new Outlook.InspectorEvents_10_PageChangeEventHandler(PageChange);

            // Initialize is called to give the derived wrappers.
            Initialize();
        }
        public static InspectorWrapper GetWrapperFor(Inspector inspector)
        {

            // Retrieve the message class by using late binding.
            string messageClass = inspector.CurrentItem.GetType().InvokeMember("MessageClass", BindingFlags.GetProperty, null, inspector.CurrentItem, null);

            // Depending on the message class, you can instantiate a
            // different wrapper explicitly for a given message class by
            // using a switch statement.
            switch (messageClass)
            {
                case "IPM.Note":
                    return new MailItemWrapper(inspector);
                case "IPM.Appointment":
                case "IPM.Appointment.test":
                    return new AppointmentItemWrapper(inspector);
            }

            // No wrapper is found.
            return null;
        }     
        protected virtual void Initialize() { }

        public void InvalidateRibbonControl(string ctrlId)
        {
            object ribbonUI = RibbonUI;

            if (ribbonUI == null || String.IsNullOrEmpty(ctrlId)) return;

            Office.IRibbonUI ribbonUIObj;
            if (ribbonUI is Array)
            {
                var ribbonUIArray = (object[])ribbonUI;
                ribbonUIObj = ribbonUIArray[0] as Office.IRibbonUI;
            }
            else
            {
                ribbonUIObj = ribbonUI as Office.IRibbonUI;
            }

            if (ribbonUIObj == null) return;
            ribbonUIObj.InvalidateControl(ctrlId);
        }

        protected virtual void Inspector_Close()
        {
            // Call the Close method - the derived classes can implement cleanup code
            // by overriding the Close method.
            Close();
            // Unregister Inspector events.
            ((Outlook.InspectorEvents_10_Event)Inspector).Close -=
                new Outlook.InspectorEvents_10_CloseEventHandler(Inspector_Close);
            ((Outlook.InspectorEvents_10_Event)Inspector).Activate -=
                new Outlook.InspectorEvents_10_ActivateEventHandler(Activate);
            ((Outlook.InspectorEvents_10_Event)Inspector).Deactivate -=
                new Outlook.InspectorEvents_10_DeactivateEventHandler(Deactivate);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMaximize -=
                new Outlook.InspectorEvents_10_BeforeMaximizeEventHandler(BeforeMaximize);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMinimize -=
                new Outlook.InspectorEvents_10_BeforeMinimizeEventHandler(BeforeMinimize);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeMove -=
                new Outlook.InspectorEvents_10_BeforeMoveEventHandler(BeforeMove);
            ((Outlook.InspectorEvents_10_Event)Inspector).BeforeSize -=
                new Outlook.InspectorEvents_10_BeforeSizeEventHandler(BeforeSize);
            ((Outlook.InspectorEvents_10_Event)Inspector).PageChange -=
                new Outlook.InspectorEvents_10_PageChangeEventHandler(PageChange);
            // Clean up resources and do a GC.Collect().
            Inspector = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            // Raise the Close event.
            if (Closed != null) Closed(Id);
        }

        protected virtual void PageChange(ref string ActivePageName) { }

        // Method is called before the inspector is resized.
        // <param name="Cancel">To prevent resizing, set Cancel to true.</param>
        protected virtual void BeforeSize(ref bool Cancel) { }

        // Method is called before the inspector is moved around.
        // <param name="Cancel">To prevent moving, set Cancel to true.</param>
        protected virtual void BeforeMove(ref bool Cancel) { }

        // Method is called before the inspector is minimized.
        // <param name="Cancel">To prevent minimizing, set Cancel to true.</param>
        protected virtual void BeforeMinimize(ref bool Cancel) { }

        // Method is called before the inspector is maximized.
        // <param name="Cancel">To prevent maximizing, set Cancel to true.</param>
        protected virtual void BeforeMaximize(ref bool Cancel) { }

        // Method is called when the inspector is deactivated.
        protected virtual void Deactivate() { }

        // Method is called when the inspector is activated.
        protected virtual void Activate() { }

        // Derived classes can do a cleanup by overriding this method.
        protected virtual void Close() { }

    }
}
