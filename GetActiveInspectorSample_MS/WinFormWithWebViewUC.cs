using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetActiveInspectorSample
{
    public partial class WinFormWithWebViewUC : Form
    {
        WebView2 wvc = null;
        public WinFormWithWebViewUC()
        {
            try
            {
                InitializeComponent();
                //CTPWebViewControl ctp = new CTPWebViewControl();
                //this.Controls.Add(ctp);

                
            }
            catch (Exception ex)
            {

            }
        }

        private void Wvc_WebMessageReceived(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
        {
            string init_model = "{\r\n  \"action\": \"INIT_MODEL\",\r\n  \"payload\": {\r\n    \"outlookId\": \"-2\",\r\n    \"mode\": \"MEETING-NEW\",\r\n    \"outgoingInd\": true,\r\n    \"isSeries\": false,\r\n    \"isOccurence\": false,\r\n    \"isExtMeetingConfWorkflow\": false\r\n  }\r\n}";
            wvc.CoreWebView2.PostWebMessageAsString(init_model);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ctpWebViewControl1.BringToFront();
            //ctpWebViewControl1.Visible = true;
            //ctpWebViewControl1.LoadHtml("test");
            InitializeWebView2Async();
        }
        private async void InitializeWebView2Async()
        {
            CoreWebView2Environment objCoreWebView2Environment = await CoreWebView2Environment.CreateAsync(null, @"c:\users\patemano\source", null);
            wvc = new WebView2();
            await wvc.EnsureCoreWebView2Async(objCoreWebView2Environment);
            wvc.Size = new Size(500, 500);
            ((ISupportInitialize)wvc).BeginInit();
            wvc.WebMessageReceived += Wvc_WebMessageReceived;
            this.Controls.Add(wvc);
            wvc.Source = new System.Uri("http://localhost.ms.com:4200/#/interactions", System.UriKind.Absolute);
            wvc.Visible = true;
            wvc.BringToFront();
            ((ISupportInitialize)wvc).EndInit();


        }
    }
}
