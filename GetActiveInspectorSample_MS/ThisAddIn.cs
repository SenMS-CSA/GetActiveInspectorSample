using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace GetActiveInspectorSample
{
    public partial class ThisAddIn
    {
        // Holds a reference to the Application.Inspectors collection.
        // Required to get notifications for NewInspector events.
        private Inspectors _inspectors;

        // A dictionary that holds a reference to the inspectors handled by the add-in.
        private Dictionary<Guid, InspectorWrapper> _wrappedInspectors;
      //  public Dictionary<object, InspectorWrapper> InspectorHandlers { get; set; }

        protected ActiveInspectorButton _activeInspectorButton;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _activeInspectorButton = new ActiveInspectorButton();

            return _activeInspectorButton;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _wrappedInspectors = new Dictionary<Guid, InspectorWrapper>();
            _inspectors = Globals.ThisAddIn.Application.Inspectors;
            _inspectors.NewInspector += new InspectorsEvents_NewInspectorEventHandler(WrapInspector);

            foreach (Inspector inspector in _inspectors)
            {
                WrapInspector(inspector);
            }
        }



        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785

            _wrappedInspectors.Clear();
            _inspectors.NewInspector -= new InspectorsEvents_NewInspectorEventHandler(WrapInspector);
            _inspectors = null;
            _activeInspectorButton.RibbonUI = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        void WrapInspector(Inspector inspector)
        {
            InspectorWrapper wrapper = InspectorWrapper.GetWrapperFor(inspector);

            wrapper.RibbonUI = _activeInspectorButton.RibbonUI;
            if (wrapper != null)
            {
                // Register the Closed event.
                wrapper.Closed += new InspectorWrapperClosedEventHandler(wrapper_Closed);
                // Remember the inspector in memory.
                _wrappedInspectors[wrapper.Id] = wrapper;

                if (wrapper.RibbonUI != null)
                {
                    wrapper.RibbonUI.InvalidateControl(ActiveInspectorButton.m_CRMBtn_ID);
                }
            }
        }

        void wrapper_Closed(Guid id)
        {
            _wrappedInspectors.Remove(id);



        }


        //public InspectorWrapper CreateNewHandler(object olInspector)
        //{
        //    InspectorWrapper inspectorWrapper = null;
        //    try
        //    {
        //        Inspector olInsp = olInspector as Inspector;
        //        if (olInsp != null && olInsp.CurrentItem != null)
        //        {
        //            if (InspectorHandlers.ContainsKey(olInsp))
        //            {
        //                inspectorWrapper = InspectorHandlers[olInsp];
        //            }
        //            else
        //            {
        //                inspectorWrapper = new InspectorWrapper(olInsp);
        //                InspectorHandlers.Add(olInsp, inspectorWrapper);
        //            }
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //    }
        //    return inspectorWrapper;
        //}

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
