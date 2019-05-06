using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Leitner_Box;
using System.Xml.Linq;

namespace Leitner_Box
{
    public partial class selectUser : Form
    {
        public selectUser()
        {
            InitializeComponent();
            listOfUsers();
        }

        private void listOfUsers()
        {
            try
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(Variables.usersFolder);
                if (!directoryInfo.Exists)
                {
                    listBox1.Items.Clear();
                    listBox1.Enabled = false;
                    buttonDelete.Enabled = false;
                    buttonSelectUser.Enabled = false;
                    return;
                }
                foreach (FileInfo file in directoryInfo.GetFiles("*.xml"))
                    listBox1.Items.Add(file.Name.Remove(file.Name.LastIndexOf('.')));
            }
            catch
            {
                listBox1.Enabled = false;
                buttonDelete.Enabled = false;
                buttonSelectUser.Enabled = false;
                listBox1.Items.Add("Exception");
            }

        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure ?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            File.Delete(Variables.usersFolder + "\\" + listBox1.SelectedItem + ".xml");
            listBox1.Items.Remove(listBox1.SelectedItem);
        }

        private void buttonSelectUser_Click(object sender, EventArgs e)
        {
            Variables.xmlFileName = Variables.usersFolder + listBox1.SelectedItem + ".xml";
            this.Close();
        }
    }
}
