using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Trendlyne_Automation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(Properties.Settings.Default.UserName) || !string.IsNullOrEmpty(Properties.Settings.Default.Password))
                backgroundWorker1.RunWorkerAsync();
            else
                MessageBox.Show("Please do the configuration for Application");
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            button1.Visible = false;
            label1.Visible = false;
            button3.Text = "Edit";
            button3.Visible = true;
            button2.Text = "Back";
            button2.Visible = true;
            label2.Text = "UserName";
            label3.Text="Password";
            label2.Visible = true;
            label3.Visible = true;
            textBox1.Text = Properties.Settings.Default.UserName;
            textBox2.Text = Properties.Settings.Default.Password;
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox1.Visible = true;
            textBox2.Visible = true;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(button3.Text=="Edit")
            {
                textBox1.Enabled = true;
                textBox2.Enabled = true;
                button3.Text = "Submit";
                button2.Text = "Cancel";
            }
            else if(button3.Text=="Submit")
            {
                Properties.Settings.Default.UserName = textBox1.Text;
                Properties.Settings.Default.Password = textBox2.Text;
                Properties.Settings.Default.Save();
                Application.Restart();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
                Application.Restart();            
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Business_Execution business_Execution = new Business_Execution();
            business_Execution.Execute();
        }
    }
}
