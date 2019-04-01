using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Kino_Simulator
{
    public partial class PerDay : Form
    {
        public PerDay()
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

           //populate list1
                string date = today.AddDays(-7).ToString(format);
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

            //populate list2
            listView2.View = View.Details;
            listView2.Columns.Add("Φορές που κληρώθηκε", 150);
            listView2.FullRowSelect = true;
            string[] arr2 = new string[2];

            List<string> stringArray2 = new List<string>();
            List<string> allNums2 = new List<string>();
            ListViewItem result2 = new ListViewItem();


            string date2 = today.AddDays(-6).ToString(format);
            string url2 = @"http://applications.opap.gr/DrawsRestServices/kino/drawDate/" + date2 + ".json";
            try
            {
                using (WebClient client = new WebClient())
                {
                    string json = client.DownloadString(url2);
                    var table = JsonConvert.DeserializeObject<RootObject>(json);

                    foreach (var item in table.draws.draw)
                    {
                        List<int> results = new List<int>();
                        results = (item.results.ToList());


                        result2.Text = item.results[0].ToString() + "," + item.results[1].ToString() + "," + item.results[2].ToString() + "," + item.results[3].ToString() + "," +
                        item.results[4].ToString() + "," + item.results[5].ToString() + "," + item.results[6].ToString() + "," + item.results[7].ToString() + "," +
                        item.results[8].ToString() + "," + item.results[9].ToString() + "," + item.results[10].ToString() + "," + item.results[11].ToString() + "," + item.results[12].ToString() + "," +
                        item.results[13].ToString() + "," + item.results[14].ToString() + "," + item.results[15].ToString() + "," + item.results[16].ToString() + "," + item.results[17].ToString() + "," +
                        item.results[18].ToString() + "," + item.results[19].ToString();

                        stringArray2.Add(result2.Text);
                    }

                    for (int j = 0; j < table.draws.draw.Count; j++)
                    {

                        var times = stringArray2[0].Count(x => x == ',');
                        for (int i = 0; i <= times; i++)
                        {
                            allNums2.Add(stringArray2[j].Split(',')[i]);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ooops... Κάτι πήγε στραβά!");
            }


            allNums2.Sort();
            var g2 = allNums2.GroupBy(i => i);

            foreach (var grp in g2)
            {
                ListViewItem itm;
                arr2[0] = grp.Key;
                arr2[1] = grp.Count().ToString();
                itm = new ListViewItem(arr2);
                listView2.Items.Add(itm);
                //Console.WriteLine("{0} {1}", grp.Key, grp.Count());
            }

            //populate list3
            listView3.View = View.Details;
            listView3.Columns.Add("Φορές που κληρώθηκε", 150);
            listView3.FullRowSelect = true;
            string[] arr3 = new string[2];

            List<string> stringArray3 = new List<string>();
            List<string> allNums3 = new List<string>();
            ListViewItem result3 = new ListViewItem();


            string date3 = today.AddDays(-5).ToString(format);
            string url3 = @"http://applications.opap.gr/DrawsRestServices/kino/drawDate/" + date3 + ".json";
            try
            {
                using (WebClient client = new WebClient())
                {
                    string json = client.DownloadString(url3);
                    var table = JsonConvert.DeserializeObject<RootObject>(json);

                    foreach (var item in table.draws.draw)
                    {
                        List<int> results = new List<int>();
                        results = (item.results.ToList());


                        result3.Text = item.results[0].ToString() + "," + item.results[1].ToString() + "," + item.results[2].ToString() + "," + item.results[3].ToString() + "," +
                        item.results[4].ToString() + "," + item.results[5].ToString() + "," + item.results[6].ToString() + "," + item.results[7].ToString() + "," +
                        item.results[8].ToString() + "," + item.results[9].ToString() + "," + item.results[10].ToString() + "," + item.results[11].ToString() + "," + item.results[12].ToString() + "," +
                        item.results[13].ToString() + "," + item.results[14].ToString() + "," + item.results[15].ToString() + "," + item.results[16].ToString() + "," + item.results[17].ToString() + "," +
                        item.results[18].ToString() + "," + item.results[19].ToString();

                        stringArray3.Add(result3.Text);
                    }

                    for (int j = 0; j < table.draws.draw.Count; j++)
                    {

                        var times = stringArray3[0].Count(x => x == ',');
                        for (int i = 0; i <= times; i++)
                        {
                            allNums3.Add(stringArray3[j].Split(',')[i]);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ooops... Κάτι πήγε στραβά!");
            }


            allNums3.Sort();
            var g3 = allNums3.GroupBy(i => i);

            foreach (var grp in g3)
            {
                ListViewItem itm;
                arr3[0] = grp.Key;
                arr3[1] = grp.Count().ToString();
                itm = new ListViewItem(arr3);
                listView3.Items.Add(itm);
                //Console.WriteLine("{0} {1}", grp.Key, grp.Count());
            }

            //populate list4
            listView4.View = View.Details;
            listView4.Columns.Add("Φορές που κληρώθηκε", 150);
            listView4.FullRowSelect = true;
            string[] arr4 = new string[2];

            List<string> stringArray4 = new List<string>();
            List<string> allNums4 = new List<string>();
            ListViewItem result4 = new ListViewItem();


            string date4 = today.AddDays(-4).ToString(format);
            string url4 = @"http://applications.opap.gr/DrawsRestServices/kino/drawDate/" + date4 + ".json";
            try
            {
                using (WebClient client = new WebClient())
                {
                    string json = client.DownloadString(url4);
                    var table = JsonConvert.DeserializeObject<RootObject>(json);

                    foreach (var item in table.draws.draw)
                    {
                        List<int> results = new List<int>();
                        results = (item.results.ToList());


                        result4.Text = item.results[0].ToString() + "," + item.results[1].ToString() + "," + item.results[2].ToString() + "," + item.results[3].ToString() + "," +
                        item.results[4].ToString() + "," + item.results[5].ToString() + "," + item.results[6].ToString() + "," + item.results[7].ToString() + "," +
                        item.results[8].ToString() + "," + item.results[9].ToString() + "," + item.results[10].ToString() + "," + item.results[11].ToString() + "," + item.results[12].ToString() + "," +
                        item.results[13].ToString() + "," + item.results[14].ToString() + "," + item.results[15].ToString() + "," + item.results[16].ToString() + "," + item.results[17].ToString() + "," +
                        item.results[18].ToString() + "," + item.results[19].ToString();

                        stringArray4.Add(result4.Text);
                    }

                    for (int j = 0; j < table.draws.draw.Count; j++)
                    {

                        var times = stringArray4[0].Count(x => x == ',');
                        for (int i = 0; i <= times; i++)
                        {
                            allNums4.Add(stringArray4[j].Split(',')[i]);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ooops... Κάτι πήγε στραβά!");
            }


            allNums4.Sort();
            var g4 = allNums4.GroupBy(i => i);

            foreach (var grp in g4)
            {
                ListViewItem itm;
                arr4[0] = grp.Key;
                arr4[1] = grp.Count().ToString();
                itm = new ListViewItem(arr4);
                listView4.Items.Add(itm);
                //Console.WriteLine("{0} {1}", grp.Key, grp.Count());
            }


            //populate list5
            listView5.View = View.Details;
            listView5.Columns.Add("Φορές που κληρώθηκε", 150);
            listView5.FullRowSelect = true;
            string[] arr5 = new string[2];

            List<string> stringArray5 = new List<string>();
            List<string> allNums5 = new List<string>();
            ListViewItem result5 = new ListViewItem();


            string date5 = today.AddDays(-3).ToString(format);
            string url5 = @"http://applications.opap.gr/DrawsRestServices/kino/drawDate/" + date5 + ".json";
            try
            {
                using (WebClient client = new WebClient())
                {
                    string json = client.DownloadString(url5);
                    var table = JsonConvert.DeserializeObject<RootObject>(json);

                    foreach (var item in table.draws.draw)
                    {
                        List<int> results = new List<int>();
                        results = (item.results.ToList());


                        result5.Text = item.results[0].ToString() + "," + item.results[1].ToString() + "," + item.results[2].ToString() + "," + item.results[3].ToString() + "," +
                        item.results[4].ToString() + "," + item.results[5].ToString() + "," + item.results[6].ToString() + "," + item.results[7].ToString() + "," +
                        item.results[8].ToString() + "," + item.results[9].ToString() + "," + item.results[10].ToString() + "," + item.results[11].ToString() + "," + item.results[12].ToString() + "," +
                        item.results[13].ToString() + "," + item.results[14].ToString() + "," + item.results[15].ToString() + "," + item.results[16].ToString() + "," + item.results[17].ToString() + "," +
                        item.results[18].ToString() + "," + item.results[19].ToString();

                        stringArray5.Add(result5.Text);
                    }

                    for (int j = 0; j < table.draws.draw.Count; j++)
                    {

                        var times = stringArray5[0].Count(x => x == ',');
                        for (int i = 0; i <= times; i++)
                        {
                            allNums5.Add(stringArray5[j].Split(',')[i]);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ooops... Κάτι πήγε στραβά!");
            }


            allNums5.Sort();
            var g5 = allNums5.GroupBy(i => i);

            foreach (var grp in g5)
            {
                ListViewItem itm;
                arr5[0] = grp.Key;
                arr5[1] = grp.Count().ToString();
                itm = new ListViewItem(arr5);
                listView5.Items.Add(itm);
                //Console.WriteLine("{0} {1}", grp.Key, grp.Count());
            }

            //populate list6
            listView6.View = View.Details;
            listView6.Columns.Add("Φορές που κληρώθηκε", 150);
            listView6.FullRowSelect = true;
            string[] arr6 = new string[2];

            List<string> stringArray6 = new List<string>();
            List<string> allNums6 = new List<string>();
            ListViewItem result6 = new ListViewItem();


            string date6 = today.AddDays(-2).ToString(format);
            string url6 = @"http://applications.opap.gr/DrawsRestServices/kino/drawDate/" + date6 + ".json";
            try
            {
                using (WebClient client = new WebClient())
                {
                    string json = client.DownloadString(url6);
                    var table = JsonConvert.DeserializeObject<RootObject>(json);

                    foreach (var item in table.draws.draw)
                    {
                        List<int> results = new List<int>();
                        results = (item.results.ToList());


                        result6.Text = item.results[0].ToString() + "," + item.results[1].ToString() + "," + item.results[2].ToString() + "," + item.results[3].ToString() + "," +
                        item.results[4].ToString() + "," + item.results[5].ToString() + "," + item.results[6].ToString() + "," + item.results[7].ToString() + "," +
                        item.results[8].ToString() + "," + item.results[9].ToString() + "," + item.results[10].ToString() + "," + item.results[11].ToString() + "," + item.results[12].ToString() + "," +
                        item.results[13].ToString() + "," + item.results[14].ToString() + "," + item.results[15].ToString() + "," + item.results[16].ToString() + "," + item.results[17].ToString() + "," +
                        item.results[18].ToString() + "," + item.results[19].ToString();

                        stringArray6.Add(result6.Text);
                    }

                    for (int j = 0; j < table.draws.draw.Count; j++)
                    {

                        var times = stringArray6[0].Count(x => x == ',');
                        for (int i = 0; i <= times; i++)
                        {
                            allNums6.Add(stringArray6[j].Split(',')[i]);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ooops... Κάτι πήγε στραβά!");
            }


            allNums6.Sort();
            var g6 = allNums6.GroupBy(i => i);

            foreach (var grp in g6)
            {
                ListViewItem itm;
                arr6[0] = grp.Key;
                arr6[1] = grp.Count().ToString();
                itm = new ListViewItem(arr6);
                listView6.Items.Add(itm);
                //Console.WriteLine("{0} {1}", grp.Key, grp.Count());
            }

            //populate list7
            listView7.View = View.Details;
            listView7.Columns.Add("Φορές που κληρώθηκε", 150);
            listView7.FullRowSelect = true;
            string[] arr7 = new string[2];

            List<string> stringArray7 = new List<string>();
            List<string> allNums7 = new List<string>();
            ListViewItem result7 = new ListViewItem();


            string date7 = today.AddDays(-1).ToString(format);
            string url7 = @"http://applications.opap.gr/DrawsRestServices/kino/drawDate/" + date7 + ".json";
            try
            {
                using (WebClient client = new WebClient())
                {
                    string json = client.DownloadString(url7);
                    var table = JsonConvert.DeserializeObject<RootObject>(json);

                    foreach (var item in table.draws.draw)
                    {
                        List<int> results = new List<int>();
                        results = (item.results.ToList());


                        result7.Text = item.results[0].ToString() + "," + item.results[1].ToString() + "," + item.results[2].ToString() + "," + item.results[3].ToString() + "," +
                        item.results[4].ToString() + "," + item.results[5].ToString() + "," + item.results[6].ToString() + "," + item.results[7].ToString() + "," +
                        item.results[8].ToString() + "," + item.results[9].ToString() + "," + item.results[10].ToString() + "," + item.results[11].ToString() + "," + item.results[12].ToString() + "," +
                        item.results[13].ToString() + "," + item.results[14].ToString() + "," + item.results[15].ToString() + "," + item.results[16].ToString() + "," + item.results[17].ToString() + "," +
                        item.results[18].ToString() + "," + item.results[19].ToString();

                        stringArray7.Add(result7.Text);
                    }

                    for (int j = 0; j < table.draws.draw.Count; j++)
                    {

                        var times = stringArray7[0].Count(x => x == ',');
                        for (int i = 0; i <= times; i++)
                        {
                            allNums7.Add(stringArray7[j].Split(',')[i]);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ooops... Κάτι πήγε στραβά!");
            }


            allNums7.Sort();
            var g7 = allNums7.GroupBy(i => i);

            foreach (var grp in g7)
            {
                ListViewItem itm;
                arr7[0] = grp.Key;
                arr7[1] = grp.Count().ToString();
                itm = new ListViewItem(arr7);
                listView7.Items.Add(itm);
                //Console.WriteLine("{0} {1}", grp.Key, grp.Count());
            }

            //populate list8
            listView8.View = View.Details;
            listView8.Columns.Add("Φορές που κληρώθηκε", 150);
            listView8.FullRowSelect = true;
            string[] arr8 = new string[2];

            List<string> stringArray8 = new List<string>();
            List<string> allNums8 = new List<string>();
            ListViewItem result8 = new ListViewItem();


            string date8 = today.ToString(format);
            string url8 = @"http://applications.opap.gr/DrawsRestServices/kino/drawDate/" + date8 + ".json";
            try
            {
                using (WebClient client = new WebClient())
                {
                    string json = client.DownloadString(url8);
                    var table = JsonConvert.DeserializeObject<RootObject>(json);

                    foreach (var item in table.draws.draw)
                    {
                        List<int> results = new List<int>();
                        results = (item.results.ToList());


                        result8.Text = item.results[0].ToString() + "," + item.results[1].ToString() + "," + item.results[2].ToString() + "," + item.results[3].ToString() + "," +
                        item.results[4].ToString() + "," + item.results[5].ToString() + "," + item.results[6].ToString() + "," + item.results[7].ToString() + "," +
                        item.results[8].ToString() + "," + item.results[9].ToString() + "," + item.results[10].ToString() + "," + item.results[11].ToString() + "," + item.results[12].ToString() + "," +
                        item.results[13].ToString() + "," + item.results[14].ToString() + "," + item.results[15].ToString() + "," + item.results[16].ToString() + "," + item.results[17].ToString() + "," +
                        item.results[18].ToString() + "," + item.results[19].ToString();

                        stringArray8.Add(result8.Text);
                    }

                    for (int j = 0; j < table.draws.draw.Count; j++)
                    {

                        var times = stringArray8[0].Count(x => x == ',');
                        for (int i = 0; i <= times; i++)
                        {
                            allNums8.Add(stringArray8[j].Split(',')[i]);
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ooops... Κάτι πήγε στραβά!");
            }


            allNums8.Sort();
            var g8 = allNums8.GroupBy(i => i);

            foreach (var grp in g8)
            {
                ListViewItem itm;
                arr8[0] = grp.Key;
                arr8[1] = grp.Count().ToString();
                itm = new ListViewItem(arr8);
                listView8.Items.Add(itm);
                //Console.WriteLine("{0} {1}", grp.Key, grp.Count());
            }
        }

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Get the new sorting column.
            ColumnHeader new_sorting_column = listView1.Columns[e.Column];
            ColumnHeader new_sorting_column2 = listView2.Columns[e.Column];
            ColumnHeader new_sorting_column3 = listView3.Columns[e.Column];
            ColumnHeader new_sorting_column4 = listView4.Columns[e.Column];
            ColumnHeader new_sorting_column5 = listView5.Columns[e.Column];
            ColumnHeader new_sorting_column6 = listView6.Columns[e.Column];
            ColumnHeader new_sorting_column7 = listView7.Columns[e.Column];
            ColumnHeader new_sorting_column8 = listView8.Columns[e.Column];

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
            listView2.ListViewItemSorter = new ListViewComparer(e.Column, sort_order);
            listView3.ListViewItemSorter = new ListViewComparer(e.Column, sort_order);
            listView4.ListViewItemSorter = new ListViewComparer(e.Column, sort_order);
            listView5.ListViewItemSorter = new ListViewComparer(e.Column, sort_order);
            listView6.ListViewItemSorter = new ListViewComparer(e.Column, sort_order);
            listView7.ListViewItemSorter = new ListViewComparer(e.Column, sort_order);
            listView8.ListViewItemSorter = new ListViewComparer(e.Column, sort_order);

            // Sort.
            listView1.Sort();
            listView2.Sort();
            listView3.Sort();
            listView4.Sort();
            listView5.Sort();
            listView6.Sort();
            listView7.Sort();
            listView8.Sort();
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

        private void button1_Click(object sender, EventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (var package = new ExcelPackage())
            {
                // Add a new worksheet to the empty workbook
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Inventory");
                //Add date/time
                ws.Cells[1, 10].Value = DateTime.Now.ToString("yyyy-MM-ddTHH:mm");

                //Add the headers
                ws.Cells["A2:Z2"].Style.Font.Bold = true;
                ws.Cells["A3:Z3"].Style.Font.Bold = true;

                ws.Cells[2, 1].Value = "Πρίν 7";
                ws.Cells[2, 2].Value = "Ημέρες";
                ws.Cells[2, 4].Value = "Πρίν 6";
                ws.Cells[2, 5].Value = "Ημέρες";
                ws.Cells[2, 7].Value = "Πρίν 5";
                ws.Cells[2, 8].Value = "Ημέρες";
                ws.Cells[2, 10].Value = "Πρίν 4";
                ws.Cells[2, 11].Value = "Ημέρες";
                ws.Cells[2, 13].Value = "Πρίν 3";
                ws.Cells[2, 14].Value = "Ημέρες";
                ws.Cells[2, 16].Value = "Πρίν 2";
                ws.Cells[2, 17].Value = "Ημέρες";
                ws.Cells[2, 19].Value = "Πρίν 1";
                ws.Cells[2, 20].Value = "Ημέρα";
                ws.Cells[2, 22].Value = "Σήμερα";
                ws.Cells[2, 23].Value = "Σήμερα";
                ws.Cells[3, 1].Value = "Αριθμός";
                ws.Cells[3, 2].Value = "Φορές";
                ws.Cells[3, 4].Value = "Αριθμός";
                ws.Cells[3, 5].Value = "Φορές";
                ws.Cells[3, 7].Value = "Αριθμός";
                ws.Cells[3, 8].Value = "Φορές";
                ws.Cells[3, 10].Value = "Αριθμός";
                ws.Cells[3, 11].Value = "Φορές";
                ws.Cells[3, 13].Value = "Αριθμός";
                ws.Cells[3, 14].Value = "Φορές";
                ws.Cells[3, 16].Value = "Αριθμός";
                ws.Cells[3, 17].Value = "Φορές";
                ws.Cells[3, 19].Value = "Αριθμός";
                ws.Cells[3, 20].Value = "Φορές";
                ws.Cells[3, 22].Value = "Αριθμός";
                ws.Cells[3, 23].Value = "Φορές";

                //Add items...               
                int i = 4;
                foreach (ListViewItem item in listView1.Items)
                {
                    ws.Cells[i, 1].Value = item.SubItems[0].Text;
                    ws.Cells[i, 2].Value = item.SubItems[1].Text;
                    i++;
                }
                i = 4;
                foreach (ListViewItem item in listView2.Items)
                {
                    ws.Cells[i, 4].Value = item.SubItems[0].Text;
                    ws.Cells[i, 5].Value = item.SubItems[1].Text;
                    i++;
                }
                i = 4;
                foreach (ListViewItem item in listView3.Items)
                {
                    ws.Cells[i, 7].Value = item.SubItems[0].Text;
                    ws.Cells[i, 8].Value = item.SubItems[1].Text;
                    i++;
                }
                i = 4;
                foreach (ListViewItem item in listView4.Items)
                {
                    ws.Cells[i, 10].Value = item.SubItems[0].Text;
                    ws.Cells[i, 11].Value = item.SubItems[1].Text;
                    i++;
                }
                i = 4;
                foreach (ListViewItem item in listView5.Items)
                {
                    ws.Cells[i, 13].Value = item.SubItems[0].Text;
                    ws.Cells[i, 14].Value = item.SubItems[1].Text;
                    i++;
                }
                i = 4;
                foreach (ListViewItem item in listView6.Items)
                {
                    ws.Cells[i, 16].Value = item.SubItems[0].Text;
                    ws.Cells[i, 17].Value = item.SubItems[1].Text;
                    i++;
                }
                i = 4;
                foreach (ListViewItem item in listView7.Items)
                {
                    ws.Cells[i, 19].Value = item.SubItems[0].Text;
                    ws.Cells[i, 20].Value = item.SubItems[1].Text;
                    i++;
                }
                i = 4;
                foreach (ListViewItem item in listView8.Items)
                {
                    ws.Cells[i, 22].Value = item.SubItems[0].Text;
                    ws.Cells[i, 23].Value = item.SubItems[1].Text;
                    i++;
                }
                var xlFile = Utils.GetFileInfo("perDay.xlsx");
                // save our new workbook in the output directory and we are done!
                package.SaveAs(xlFile);
                MessageBox.Show("Αποθηκεύτηκε στην Επιφάνεια Εργασίας");

                ////if you have Microsoft Office Excel use the code below
                //using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel|*.xls", ValidateNames = true })
                //{
                //    if (sfd.ShowDialog() == DialogResult.OK)
                //    {
                //        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                //        Workbook wb = app.Workbooks.Add(XlSheetType.xlWorksheet);
                //        Worksheet ws = (Worksheet)app.ActiveSheet;
                //        app.Visible = false;
                //        ws.Cells[1, 1] = "Αριθμός";
                //        ws.Cells[1, 2] = "Φορές που εμφανίστηκε";
                //        ws.Cells[1, 3] = "Αριθμός";
                //        ws.Cells[1, 4] = "Φορές που εμφανίστηκε";
                //        ws.Cells[1, 5] = "Αριθμός";
                //        ws.Cells[1, 6] = "Φορές που εμφανίστηκε";
                //        ws.Cells[1, 7] = "Αριθμός";
                //        ws.Cells[1, 8] = "Φορές που εμφανίστηκε";
                //        ws.Cells[1, 9] = "Αριθμός";
                //        ws.Cells[1, 10] = "Φορές που εμφανίστηκε";
                //        ws.Cells[1, 11] = "Αριθμός";
                //        ws.Cells[1, 12] = "Φορές που εμφανίστηκε";
                //        ws.Cells[1, 13] = "Αριθμός";
                //        ws.Cells[1, 14] = "Φορές που εμφανίστηκε";
                //        ws.Cells[1, 15] = "Αριθμός";
                //        ws.Cells[1, 16] = "Φορές που εμφανίστηκε";
                //        int i = 2;
                //        foreach (ListViewItem item in listView1.Items)
                //        {
                //            ws.Cells[i, 1] = item.SubItems[0].Text;
                //            ws.Cells[i, 2] = item.SubItems[1].Text;
                //            i++;
                //        }
                //        i = 2;
                //        foreach (ListViewItem item in listView2.Items)
                //        {
                //            ws.Cells[i, 3] = item.SubItems[0].Text;
                //            ws.Cells[i, 4] = item.SubItems[1].Text;
                //            i++;
                //        }
                //        i = 2;
                //        foreach (ListViewItem item in listView3.Items)
                //        {
                //            ws.Cells[i, 5] = item.SubItems[0].Text;
                //            ws.Cells[i, 6] = item.SubItems[1].Text;
                //            i++;
                //        }
                //        i = 2;
                //        foreach (ListViewItem item in listView4.Items)
                //        {
                //            ws.Cells[i, 7] = item.SubItems[0].Text;
                //            ws.Cells[i, 8] = item.SubItems[1].Text;
                //            i++;
                //        }
                //        i = 2;
                //        foreach (ListViewItem item in listView5.Items)
                //        {
                //            ws.Cells[i, 9] = item.SubItems[0].Text;
                //            ws.Cells[i, 10] = item.SubItems[1].Text;
                //            i++;
                //        }
                //        i = 2;
                //        foreach (ListViewItem item in listView6.Items)
                //        {
                //            ws.Cells[i, 11] = item.SubItems[0].Text;
                //            ws.Cells[i, 12] = item.SubItems[1].Text;
                //            i++;
                //        }
                //        i = 2;
                //        foreach (ListViewItem item in listView7.Items)
                //        {
                //            ws.Cells[i, 13] = item.SubItems[0].Text;
                //            ws.Cells[i, 14] = item.SubItems[1].Text;
                //            i++;
                //        }
                //        i = 2;
                //        foreach (ListViewItem item in listView8.Items)
                //        {
                //            ws.Cells[i, 15] = item.SubItems[0].Text;
                //            ws.Cells[i, 16] = item.SubItems[1].Text;
                //            i++;
                //        }
                //        wb.SaveAs(sfd.FileName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                //        app.Quit();
                //        MessageBox.Show("Αποθηκεύτηκαν σε Excel");
                //    }
            }

        }
    }
}
