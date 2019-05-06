using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Leitner_Box
{
    public partial class aboutMe : Form
    {
        public aboutMe()
        {
            InitializeComponent();
            linkLabel1.Links[0].LinkData = "http://www.mds-soft.persianblog.ir/";
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(e.Link.LinkData.ToString());
            }
            catch { }
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Left)
                    System.Diagnostics.Process.Start("http://www.mds-soft.persianblog.ir/");
            }
            catch { }
        }
    }
}
