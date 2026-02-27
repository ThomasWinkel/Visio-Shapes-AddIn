using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using Newtonsoft.Json;
using System.Text;
using System.Net.Http.Headers;

namespace VisioAddin.Ui
{
    public partial class FrmContributeShape : Form
    {
        private Visio.Shape Shape;
        private Visio.Master Master;
        private static readonly HttpClient client = new HttpClient();

        public FrmContributeShape(Visio.Shape shape)
        {
            InitializeComponent();
            Shape = shape;
            Master = shape.Master;

            if (Master != null)
            {
                tbName.Text = Master.Name;
                tbPrompt.Text = Master.Prompt;
                tbKeywords.Text = Master.PageSheet.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowMisc, (short)Visio.VisCellIndices.visObjKeywords].ResultStr[""];
            }

            btnContribute.Enabled = false;
            cbTeam.SelectedIndexChanged += cbTeam_SelectedIndexChanged;
            this.Load += async (s, e) => await LoadTeamsAsync();
        }

        private async Task LoadTeamsAsync()
        {
            var items = new List<Models.TeamItem>
            {
                new Models.TeamItem { Id = null, Name = "Persönlich" }
            };

            try
            {
                string token = Globals.ThisAddIn.ServerHandler.CurrentServerToken;
                string url = Globals.ThisAddIn.ServerHandler.CurrentServerUrl + "/get_user_teams";

                using (var request = new HttpRequestMessage(HttpMethod.Get, url))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    var response = await client.SendAsync(request).ConfigureAwait(false);
                    string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    var teams = JsonConvert.DeserializeObject<List<Models.TeamItem>>(json);
                    if (teams != null)
                        items.AddRange(teams);
                }
            }
            catch { }

            cbTeam.Invoke(new Action(() =>
            {
                cbTeam.Items.Clear();
                foreach (var item in items)
                    cbTeam.Items.Add(item);
            }));
        }

        private void cbTeam_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnContribute.Enabled = cbTeam.SelectedIndex >= 0;
        }

        public async Task Submit()
        {
            if (tbName.Text == "")
            {
                MessageBox.Show("No name defined.", "Contribute Shape");
                return;
            }

            List<string> formats = new List<string>
            {
                "Visio 15.0 Masters",
                "Object Descriptor"
            };

            Visio.Document docStencil = null;

            if (Master == null)
            {
                docStencil = Globals.ThisAddIn.Application.Documents.AddEx("", Visio.VisMeasurementSystem.visMSDefault, (short)Visio.VisOpenSaveArgs.visAddStencil + (short)Visio.VisOpenSaveArgs.visAddHidden);
                Shape.Copy();
                docStencil.Masters.Paste();
                Master = docStencil.Masters[1];
                Master.Name = tbName.Text;
                Master.Prompt = tbPrompt.Text;
                Master.PageSheet.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowMisc, (short)Visio.VisCellIndices.visObjKeywords].Formula = "\"" + tbKeywords.Text + "\"";
            }

            DataObject dataObject = new DataObject(Master);
            Dictionary<string, string> dictDataObject = new Dictionary<string, string>();
            
            foreach (string format in dataObject.GetFormats(false))
            {
                dictDataObject[format] = null;
                MemoryStream streamData = (MemoryStream)dataObject.GetData(format);

                if (streamData != null && formats.Contains(format))
                {
                    dictDataObject[format] = Convert.ToBase64String(streamData.ToArray());
                }
            }

            string imagePath = Path.Combine(Path.GetTempPath(), "image.png");
            Master.Export(imagePath);
            Image image = Image.FromFile(imagePath);
            byte[] paramFileStream = ImageToByte2(image);
            image.Dispose();
            File.Delete(imagePath);

            Models.OnlineShape onlineShape = new Models.OnlineShape();

            onlineShape.Name = tbName.Text;
            onlineShape.Prompt = tbPrompt.Text;
            onlineShape.Keywords = tbKeywords.Text;
            onlineShape.DataObject = JsonConvert.SerializeObject(dictDataObject);
            onlineShape.TeamId = (cbTeam.SelectedItem as Models.TeamItem)?.Id;

            string json = JsonConvert.SerializeObject(onlineShape);

            using (var content = new MultipartFormDataContent())
            {
                content.Add(new StringContent(json, Encoding.UTF8, "application/json"), "json");
                content.Add(new StreamContent(new MemoryStream(paramFileStream)), "image", "image.png");

                // Use singleton HttpClient with per-request headers
                string token = Globals.ThisAddIn.ServerHandler.CurrentServerToken;
                if (Globals.ThisAddIn.ServerHandler.CurrentServerUrl.StartsWith("https"))
                {
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                }

                using (var request = new HttpRequestMessage(HttpMethod.Post, Globals.ThisAddIn.ServerHandler.CurrentServerUrl + "/add_shape"))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    request.Content = content;

                    var response = await client.SendAsync(request).ConfigureAwait(false);
                    string strResponse = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    // Return to UI thread for COM operations and MessageBox
                    if (InvokeRequired)
                    {
                        Invoke(new Action(() => CloseStencilAndShowResult(docStencil, strResponse)));
                    }
                    else
                    {
                        CloseStencilAndShowResult(docStencil, strResponse);
                    }
                }
            }
        }

        private void CloseStencilAndShowResult(Visio.Document docStencil, string strResponse)
        {
            // COM operations must be on UI thread
            if (docStencil != null)
            {
                docStencil.Saved = true;
                docStencil.Close();
            }

            MessageBox.Show(strResponse);
            if (strResponse == "Failed")
            {
                MessageBox.Show("Upload failed, check credentials.");
                return;
            }

            if (Globals.ThisAddIn._panelManager.IsPanelOpened(Globals.ThisAddIn.Application.ActiveWindow))
            {
                Globals.ThisAddIn.RaiseEventOnContribute(tbName.Text);
            }
        }

        public static byte[] ImageToByte2(Image img)
        {
            byte[] byteArray = new byte[0];
            using (MemoryStream stream = new MemoryStream())
            {
                img.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                stream.Close();

                byteArray = stream.ToArray();
            }
            return byteArray;
        }

        private async void btnContribute_Click(object sender, EventArgs e)
        {
            try
            {
                await Submit().ConfigureAwait(false);
                Invoke(new Action(() => this.Close()));
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                Invoke(new Action(() => MessageBox.Show($"Failed to submit shape: {msg}", "Upload Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error)));
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}