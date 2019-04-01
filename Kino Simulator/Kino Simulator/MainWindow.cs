﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kino_Simulator
{
    public partial class MainWindow : Form
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Search search = new Search();
            search.Show();
            this.Hide();    
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Popular popular = new Popular();
            popular.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PerDay perDay = new PerDay();
            perDay.Show();
            this.Hide();
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {

        }
    }
}
