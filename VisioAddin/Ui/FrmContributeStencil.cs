using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using System.Text;
using VisioAddin.Models;
using System.Runtime.InteropServices.ComTypes;

namespace VisioAddin.Ui
{
    public partial class FrmContributeStencil : Form
    {
        private static readonly HttpClient client = new HttpClient();

        public FrmContributeStencil()
        {
            InitializeComponent();

            foreach (Visio.Document doc in Globals.ThisAddIn.Application.Documents)
            {
                if (doc.Type == Visio.VisDocumentTypes.visTypeStencil)
                {
                    lbStencils.Items.Add(doc.Name);
                }
            }

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
            UpdateContributeButton();
        }

        private async void btnContribute_Click(object sender, EventArgs e)
        {
            try
            {
                await SubmitStencilAsync().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                Invoke(new Action(() => MessageBox.Show($"Failed to submit stencil: {msg}", "Upload Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error)));
            }
        }

        private async Task SubmitStencilAsync()
        {
            Visio.Document stencil = null;

            foreach (Visio.Document doc in Globals.ThisAddIn.Application.Documents)
            {
                if (doc.Type == Visio.VisDocumentTypes.visTypeStencil)
                {
                    if (doc.Name == lbStencils.SelectedItem.ToString())
                    {
                        stencil = doc;
                        break;
                    }
                }
            }

            if (stencil == null) return;

            using (var content = new MultipartFormDataContent())
            using (var stream = File.OpenRead(stencil.FullName))
            {
                var streamContent = new StreamContent(stream);
                content.Add(streamContent, "stencil", stencil.Name);

                Models.OnlineStencil onlineStencil = new Models.OnlineStencil();
                onlineStencil.FileName = stencil.Name;
                onlineStencil.Title = stencil.Title;
                onlineStencil.Subject = stencil.Subject;
                onlineStencil.Author = stencil.Creator;
                onlineStencil.Manager = stencil.Manager;
                onlineStencil.Company = stencil.Company;
                onlineStencil.Language = stencil.Language.ToString();
                onlineStencil.Categories = stencil.Category;
                onlineStencil.Tags = stencil.Keywords;
                onlineStencil.Comments = stencil.Description;
                onlineStencil.TeamId = (cbTeam.SelectedItem as Models.TeamItem)?.Id;

                List<string> formats = new List<string>
                {
                    "Visio 15.0 Masters",
                    "Object Descriptor"
                };

                int i = 0;

                foreach (Visio.Master master in stencil.Masters)
                {
                    DataObject dataObject = new DataObject(master);
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

                    Models.OnlineShape onlineShape = new Models.OnlineShape();
                    onlineShape.Name = master.Name;
                    onlineShape.Prompt = master.Prompt;
                    onlineShape.Keywords = master.PageSheet.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowMisc, (short)Visio.VisCellIndices.visObjKeywords].ResultStr[""];
                    onlineShape.DataObject = JsonConvert.SerializeObject(dictDataObject);
                    onlineStencil.Shapes.Add(onlineShape);

                    string imagePath = Path.Combine(Path.GetTempPath(), "image.png");
                    master.Export(imagePath);
                    Image image = Image.FromFile(imagePath);
                    byte[] paramFileStream = ImageToByte2(image);
                    image.Dispose();
                    File.Delete(imagePath);

                    // StreamContent takes ownership of MemoryStream and disposes it
                    MemoryStream memoryStream = new MemoryStream(paramFileStream);
                    streamContent = new StreamContent(memoryStream);

                    i++;
                    content.Add(streamContent, "images", i.ToString() + ".png");
                }

                string json = JsonConvert.SerializeObject(onlineStencil);

                StringContent stringContent = new StringContent(json, Encoding.UTF8, "application/json");
                content.Add(stringContent, "json");

                // Use singleton HttpClient with per-request headers
                string token = Globals.ThisAddIn.ServerHandler.CurrentServerToken;
                if (Globals.ThisAddIn.ServerHandler.CurrentServerUrl.StartsWith("https"))
                {
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                }

                using (var request = new HttpRequestMessage(HttpMethod.Post, Globals.ThisAddIn.ServerHandler.CurrentServerUrl + "/add_stencil"))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    request.Content = content;

                    var response = await client.SendAsync(request).ConfigureAwait(false);
                    string strResponse = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    // Return to UI thread for MessageBox and CloseForm
                    if (InvokeRequired)
                    {
                        Invoke(new Action(() => ShowResultAndClose(strResponse)));
                    }
                    else
                    {
                        ShowResultAndClose(strResponse);
                    }
                }
            }
        }

        private void ShowResultAndClose(string strResponse)
        {
            MessageBox.Show(strResponse);
            if (strResponse == "Failed")
            {
                MessageBox.Show("Upload failed, check credentials.");
                return;
            }

            CloseForm();
        }

        private delegate void BlankDelegate();

        private void CloseForm()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new BlankDelegate(this.CloseForm));
            }
            else
            {
                this.Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void lbStencils_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateContributeButton();
        }

        private void UpdateContributeButton()
        {
            btnContribute.Enabled = lbStencils.SelectedItems.Count > 0 && cbTeam.SelectedIndex >= 0;
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
    }
}
