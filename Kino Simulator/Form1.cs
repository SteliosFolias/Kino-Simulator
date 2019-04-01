using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Runtime.Serialization;
using System.Net;

namespace Kino_Simulator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            //http://applications.opap.gr/DrawsRestServices/kino/drawDate/25-04-2018.json

            DateTime today = DateTime.Today;
            string format = "dd-MM-yyyy";
            string url = @"http://applications.opap.gr/DrawsRestServices/kino/drawDate/" + today.ToString(format) + ".json";
            try
            {
                using (WebClient client = new WebClient())
                {

                    string json = client.DownloadString(url);
                    var table = JsonConvert.DeserializeObject<RootObject>(json);

                    listView1.Columns.Add("Ημερ/νία και ώρα Κλήρωσης", 160);
                    listView1.Columns.Add("Αποτελέσματα", 310);
                    listView1.View = View.Details;
                    listView1.FullRowSelect = true;
                    string[] arr = new string[2];

                    foreach (var item in table.draws.draw)
                    {
                        ListViewItem result = new ListViewItem();
                        ListViewItem itm;
                        List<int> results = new List<int>();
                        results = (item.results.ToList());

                        result.Text = item.results[0].ToString() + "," + item.results[1].ToString() + "," + item.results[2].ToString() + "," + item.results[3].ToString() + "," +
                        item.results[4].ToString() + "," + item.results[5].ToString() + "," + item.results[6].ToString() + "," + item.results[7].ToString() + "," +
                         item.results[8].ToString() + "," + item.results[9].ToString() + "," + item.results[10].ToString() + "," + item.results[11].ToString() + "," + item.results[12].ToString() + "," +
                        item.results[13].ToString() + "," + item.results[14].ToString() + "," + item.results[15].ToString() + "," + item.results[16].ToString() + "," + item.results[17].ToString() + "," +
                         item.results[18].ToString() + "," + item.results[19].ToString();

                        arr[0] = item.drawTime;
                        arr[1] = result.Text;
                        itm = new ListViewItem(arr);
                        listView1.Items.Add(itm);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ooops... Κάτι πήγε στραβά!");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}