using Microsoft.Office.Interop.Outlook;
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Cancel"></param>
        void Item_Open(ref bool Cancel)
        {
            var aptItemParent = Item.Parent;
            string storeId = string.Empty;
            if (aptItemParent is MAPIFolder)
            {
                MAPIFolder folder = (MAPIFolder)aptItemParent;
                
                var owner = folder.GetOwner();

                storeId = folder.StoreID;
                
            }
            EntryID entryID = null;
            try
            {
                //Microsoft Team: Here storeId (for shared calendar) is coming either primary user's address or null/random text (observed for few of the users)
                //when "Shared calendar feture is ON"
                //If disable this feature and it works fine.
                entryID = new EntryID(storeId);
                System.Windows.Forms.MessageBox.Show($"User Address:  { entryID.UserAddress}");
            }
            finally
            {
            }
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

    #region Sub class EntryID
    ///<summary>
    /// EntryID
    /// This class is used to hold EntryID and to retrieve details like ServerShortName, and UserAddress
    /// Reference: Store Entry ID v2: http://blogs.msdn.com/b/stephen_griffin/archive/2011/07/21/store-entry-id-v2.aspx
    ///</summary>
    public class EntryID
    {
        private string entryId = "";
        private string serverShortName = "";
        private string userAddress = "";
        /// <summary>
        /// 
        /// </summary>
        /// <param name="anEntryId"></param>
        public EntryID(string anEntryId)
        {
            entryId = anEntryId;
        }
        /// <summary>
        /// 
        /// </summary>
        public string ServerShortName
        {
            get
            {
                return getServerShortName();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public string UserAddress
        {
            get
            {
                return getUserAddress();
            }

        }

        private string getServerShortName()
        {
            if (serverShortName != "") return serverShortName;

            //Fix for StoreID issue for ServerShortName where it triggered exception like "Index and length must refer to a location within the string." in HexToByteArray
            //e.g. for Imae, Akihiko (Enterprise Infrastructure), and Taylor, Tokiko (Enterprise Infrastructure) users StoreID values.
            if (entryId.IndexOf("00", 120) % 2 == 0)
            {
                byte[] temp = HexToByteArray(entryId.Substring(120, entryId.IndexOf("00", 120) - 120));
                serverShortName = Encoding.ASCII.GetString(temp);
            }
            else if (entryId.IndexOf("00", 120) % 2 == 1)
            {
                byte[] temp = HexToByteArray(entryId.Substring(120, (entryId.IndexOf("00", 120) + 1) - 120));
                serverShortName = Encoding.ASCII.GetString(temp);
            }
            return serverShortName;
        }

        private string getUserAddress()
        {
            if (userAddress != "") return userAddress;
            serverShortName = getServerShortName();
            byte[] temp = HexToByteArray(entryId.Substring(120 + serverShortName.Length * 2 + 2));
            userAddress = Encoding.ASCII.GetString(temp);
            return userAddress;
        }

        private byte[] HexToByteArray(string hex)
        {
            byte[] bytes = new byte[hex.Length / 2];

            for (int i = 0; i < hex.Length; i += 2)
            {
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            }

            return bytes;
        }
    }
    #endregion
}

