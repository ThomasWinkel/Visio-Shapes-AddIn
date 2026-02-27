using Microsoft.Web.WebView2.Core;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioAddin
{
    public partial class TheForm : Form
    {
        private readonly Visio.Window _window;
        private bool initializing = true;

        /// <summary>
        /// Form constructor, receives parent Visio diagram window
        /// </summary>
        /// <param name="window">Parent Visio diagram window</param>
        public TheForm(Visio.Window window)
        {
            _window = window;
            InitializeComponent();
            // Fire-and-forget with proper exception handling
            _ = InitializeAsyncSafe();
            Globals.ThisAddIn.OnContribute += OnContribute_Raised;
            InitializeForms();
        }

        private void InitializeForms()
        {
            initializing = true;

            cbServer.Items.Clear();

            if (Globals.ThisAddIn.ServerHandler.ServerSettings.Servers.Count > 0)
            {
                foreach (var server in Globals.ThisAddIn.ServerHandler.ServerSettings.Servers)
                {
                    cbServer.Items.Add(server.Name);
                }
                cbServer.SelectedItem = Globals.ThisAddIn.ServerHandler.CurrentServerName;
            }

            initializing = false;
        }

        private async Task InitializeAsyncSafe()
        {
            try
            {
                await InitializeAsync().ConfigureAwait(true); // Stay on UI thread
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing web view: {ex.Message}", "Initialization Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task InitializeAsync()
        {
            string userDataFolder = Path.Combine(Path.GetTempPath(), System.Reflection.Assembly.GetExecutingAssembly().GetName().Name);
            CoreWebView2Environment cwv2Environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder, null).ConfigureAwait(true);
            await webView.EnsureCoreWebView2Async(cwv2Environment).ConfigureAwait(true);
            webView.CoreWebView2.AddHostObjectToScript("WebViewDragDrop", new WebViewDragDrop(this));
            Login();
        }

        private void Login()
        {
            string postDataString = "token=" + Globals.ThisAddIn.ServerHandler.CurrentServerToken;
            UTF8Encoding utfEncoding = new UTF8Encoding();
            byte[] postData = utfEncoding.GetBytes(postDataString);
            MemoryStream postDataStream = new MemoryStream(postDataString.Length);
            postDataStream.Write(postData, 0, postData.Length);
            postDataStream.Seek(0, SeekOrigin.Begin);

            CoreWebView2WebResourceRequest webResourceRequest =
                webView.CoreWebView2.Environment.CreateWebResourceRequest(
                Globals.ThisAddIn.ServerHandler.CurrentServerUrl + "/token_login",
                "POST",
                postDataStream,
                "Content-Type: application/x-www-form-urlencoded\r\n");
            try
            {
                webView.CoreWebView2.NavigateWithWebResourceRequest(webResourceRequest);
            }
            catch { }
        }

        delegate void SearchCallback(string search);

        private void SetSearch(string search)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (webView.InvokeRequired)
            {
                SearchCallback d = new SearchCallback(SetSearch);
                this.Invoke(d, new object[] { search });
            }
            else
            {
                Search(search);
            }
        }

        private void OnContribute_Raised(object sender, EventArgs e)
        {
            SetSearch(sender as string);
        }

        private void Search(string search)
        {
            string postDataString = "search=" + search;
            UTF8Encoding utfEncoding = new UTF8Encoding();
            byte[] postData = utfEncoding.GetBytes(postDataString);
            MemoryStream postDataStream = new MemoryStream(postDataString.Length);
            postDataStream.Write(postData, 0, postData.Length);
            postDataStream.Seek(0, SeekOrigin.Begin);

            CoreWebView2WebResourceRequest webResourceRequest =
                webView.CoreWebView2.Environment.CreateWebResourceRequest(
                Globals.ThisAddIn.ServerHandler.CurrentServerUrl,
                "POST",
                postDataStream,
                "Content-Type: application/x-www-form-urlencoded\r\n");
            try
            {
                webView.CoreWebView2.NavigateWithWebResourceRequest(webResourceRequest);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Search navigation failed: {ex.Message}");
            }
        }

        // Open original preview in standard browser
        private void webView_NavigationStarting(object sender, CoreWebView2NavigationStartingEventArgs e)
        {
            if (!(e.Uri.ToString().Equals(Globals.ThisAddIn.ServerHandler.CurrentServerUrl + "/", StringComparison.InvariantCultureIgnoreCase)))
            {
                //e.Cancel = true;
                //System.Diagnostics.Process.Start(e.Uri.ToString());
            }
        }

        // Accessed from JavaScript
        [ComVisible(true)]
        public class WebViewDragDrop
        {
            readonly TheForm onlineStencilsForm;

            public WebViewDragDrop(TheForm m)
            {
                this.onlineStencilsForm = m;
            }

            public void DragDropShape(string shapeData)
            {
                try
                {
                    Dictionary<string, string> dictDataObject = JsonConvert.DeserializeObject<Dictionary<string, string>>(shapeData);
                    DataObject dataObject = new DataObject();

                    foreach (var format in dictDataObject.Keys)
                    {
                        if (dictDataObject[format] != null)
                        {
                            MemoryStream streamData = new MemoryStream(Convert.FromBase64String(dictDataObject[format]));
                            dataObject.SetData(format, streamData);
                        }
                        else
                        {
                            dataObject.SetData(format, null);
                        }

                    }

                    onlineStencilsForm.DoDragDrop(dataObject, DragDropEffects.All);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Drag-drop failed: {ex.Message}");
                    MessageBox.Show($"Failed to drag shape: {ex.Message}", "Drag Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void btnContributeShape_Click(object sender, EventArgs e)
        {
            Visio.Shape shape = null;
            if (Globals.ThisAddIn.Application.ActiveWindow.Type == (short)Visio.VisWinTypes.visDrawing)
            {
                shape = Globals.ThisAddIn.Application.ActiveWindow.Selection.PrimaryItem;
            }
            if (shape == null)
            {
                MessageBox.Show("No shape selected.", "Contribute Shape");
                return;
            }
            var form = new Ui.FrmContributeShape(shape);
            form.ShowDialog();
        }

        private void btnContributeStencil_Click(object sender, EventArgs e)
        {
            var form = new Ui.FrmContributeStencil();
            form.ShowDialog();
        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            Ui.frmSettings form = new Ui.frmSettings();
            form.ShowDialog();
            InitializeForms();
        }

        private void cbServer_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing) return;
            Globals.ThisAddIn.ServerHandler.CurrentServerName = cbServer.ComboBox.SelectedItem.ToString();
            Login();
        }
    }
}