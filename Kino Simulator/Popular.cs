using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kino_Simulator
{
    public partial class Popular : Form
    {

        public Popular()
        {
            InitializeComponent();

            DateTime today = DateTime.Today;
            string format = "dd-MM-yyyy";
            listView1.View = View.Details;
            listView1.Columns.Add("Φορές που κληρώθηκε", 150);
            listView1.FullRowSelect = true;
            string[] arr = new string[2];

            List<string> stringArray = new List<string>();
            List<string> allNums = new List<string>();
            ListViewItem result = new ListViewItem();

            for (int day = 0; day < 8; day++)
            {
                string date = today.AddDays(-day).ToString(format);
                string url = @"http://applications.opap.gr/DrawsRestServices/kino/drawDate/" + date + ".json";
                try
                {
                    using (WebClient client = new WebClient())
                    {
                        string json = client.DownloadString(url);
                        var table = JsonConvert.DeserializeObject<RootObject>(json);
                       
                        foreach (var item in table.draws.draw)
                        {
                            List<int> results = new List<int>();
                            results = (item.results.ToList());


                            result.Text = item.results[0].ToString() + "," + item.results[1].ToString() + "," + item.results[2].ToString() + "," + item.results[3].ToString() + "," +
                            item.results[4].ToString() + "," + item.results[5].ToString() + "," + item.results[6].ToString() + "," + item.results[7].ToString() + "," +
                            item.results[8].ToString() + "," + item.results[9].ToString() + "," + item.results[10].ToString() + "," + item.results[11].ToString() + "," + item.results[12].ToString() + "," +
                            item.results[13].ToString() + "," + item.results[14].ToString() + "," + item.results[15].ToString() + "," + item.results[16].ToString() + "," + item.results[17].ToString() + "," +
                            item.results[18].ToString() + "," + item.results[19].ToString();

                            stringArray.Add(result.Text);
                        }

                        for (int j = 0; j < table.draws.draw.Count; j++)
                        {

                            var times = stringArray[0].Count(x => x == ',');
                            for (int i = 0; i <= times; i++)
                            {
                                allNums.Add(stringArray[j].Split(',')[i]);
                            }
                        }                     
                    }
                }
                catch
                {
                    MessageBox.Show("Ooops... Κάτι πήγε στραβά!");
                }
            }

            allNums.Sort();
            var g = allNums.GroupBy(i => i);

            foreach (var grp in g)
            {
                ListViewItem itm;
                arr[0] = grp.Key;
                arr[1] = grp.Count().ToString();
                itm = new ListViewItem(arr);
                listView1.Items.Add(itm);
                //Console.WriteLine("{0} {1}", grp.Key, grp.Count());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Get the new sorting column.
            ColumnHeader new_sorting_column = listView1.Columns[e.Column];

            // Figure out the new sorting order.
            System.Windows.Forms.SortOrder sort_order;
            if (Position == null)
            {
                // New column. Sort ascending.
                sort_order = SortOrder.Ascending;
            }
            else
            {
                // See if this is the same column.
                if (new_sorting_column == Position)
                {
                    // Same column. Switch the sort order.
                    if (Position.Text.StartsWith("> "))
                    {
                        sort_order = SortOrder.Descending;
                    }
                    else
                    {
                        sort_order = SortOrder.Ascending;
                    }
                }
                else
                {
                    // New column. Sort ascending.
                    sort_order = SortOrder.Ascending;
                }

                // Remove the old sort indicator.
                Position.Text = "Αριθμός";
            }

            // Display the new sort order.
            Position = new_sorting_column;
            if (sort_order == SortOrder.Ascending)
            {
                Position.Text = "> " + Position.Text;
            }
            else
            {
                Position.Text = "< " + Position.Text;
            }

            // Create a comparer.
            listView1.ListViewItemSorter = new ListViewComparer(e.Column, sort_order);

            // Sort.
            listView1.Sort();
        }

        public class ListViewComparer : System.Collections.IComparer
        {
            private int ColumnNumber;
            private SortOrder SortOrder;

            public ListViewComparer(int column_number,
                SortOrder sort_order)
            {
                ColumnNumber = column_number;
                SortOrder = sort_order;
            }

            // Compare two ListViewItems.
            public int Compare(object object_x, object object_y)
            {
                // Get the objects as ListViewItems.
                ListViewItem item_x = object_x as ListViewItem;
                ListViewItem item_y = object_y as ListViewItem;

                // Get the corresponding sub-item values.
                string string_x;
                if (item_x.SubItems.Count <= ColumnNumber)
                {
                    string_x = "";
                }
                else
                {
                    string_x = item_x.SubItems[ColumnNumber].Text;
                }

                string string_y;
                if (item_y.SubItems.Count <= ColumnNumber)
                {
                    string_y = "";
                }
                else
                {
                    string_y = item_y.SubItems[ColumnNumber].Text;
                }

                // Compare them.
                int result;
                double double_x, double_y;
                if (double.TryParse(string_x, out double_x) &&
                    double.TryParse(string_y, out double_y))
                {
                    // Treat as a number.
                    result = double_x.CompareTo(double_y);
                }
                else
                {
                    DateTime date_x, date_y;
                    if (DateTime.TryParse(string_x, out date_x) &&
                        DateTime.TryParse(string_y, out date_y))
                    {
                        // Treat as a date.
                        result = date_x.CompareTo(date_y);
                    }
                    else
                    {
                        // Treat as a string.
                        result = string_x.CompareTo(string_y);
                    }
                }

                // Return the correct result depending on whether
                // we're sorting ascending or descending.
                if (SortOrder == SortOrder.Ascending)
                {
                    return result;
                }
                else
                {
                    return -result;
                }
            }
        }

        private void Popular_Load(object sender, EventArgs e)
        {

        }
    }
}
