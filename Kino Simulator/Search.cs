using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Runtime.Serialization;
using System.Net;

namespace Kino_Simulator
{
    public partial class Search : Form
    {
        public Search()
        {
            InitializeComponent();
            listView1.Columns.Add("Αποτελέσματα", 310);
            listView1.View = View.Details;
        }

       
        private void button1_Click(object sender, EventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string url = @"http://applications.opap.gr/DrawsRestServices/kino/" + textBox1.Text + ".json";
            try
            {
                using (WebClient client = new WebClient())
                {

                    string json = client.DownloadString(url);
                    var table = JsonConvert.DeserializeObject<RootObject>(json);
       
                        List<int> results = new List<int>();
                        results = table.draw.results;

                        ListViewItem result = new ListViewItem();

                        result.Text = results[0].ToString() + "," + results[1].ToString() + "," + results[2].ToString() + "," + results[3].ToString() + "," +
                        results[4].ToString() + "," + results[5].ToString() + "," + results[6].ToString() + "," + results[7].ToString() + "," +
                         results[8].ToString() + "," + results[9].ToString() + "," + results[10].ToString() + "," + results[11].ToString() + "," + results[12].ToString() + "," +
                        results[13].ToString() + "," + results[14].ToString() + "," + results[15].ToString() + "," + results[16].ToString() + "," + results[17].ToString() + "," +
                         results[18].ToString() + "," + results[19].ToString();

                        listView1.Items.Add(result);            
                }
            }
            catch
            {
                MessageBox.Show("Δεν υπάρχει κλήρωση με αριθμό " + textBox1.Text);
            }
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "Αριθμός Κλήρωσης")
            {
                textBox1.Text = "";
            }
            else
            {
                textBox1.Text = textBox1.Text;
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = "Αριθμός Κλήρωσης";
            }
            else
            {
                textBox1.Text = textBox1.Text;
            }
        }

        private void Search_Load(object sender, EventArgs e)
        {

        }
    }
}
