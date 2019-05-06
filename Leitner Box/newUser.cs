using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using System.IO;
using Leitner_Box;

namespace Leitner_Box
{
    public partial class newUser : Form
    {
        public newUser()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Variables.xDocument = XDocument.Load(Application.StartupPath + "\\" + "Default.XML");

                var startTime = (from q in Variables.xDocument.Descendants("StartTime") select q).First();
                startTime.Value = DateTime.Now.ToString();

                if (!Directory.Exists(Variables.usersFolder)) Directory.CreateDirectory(Variables.usersFolder);
                if (File.Exists(Variables.usersFolder + "\\" + maskedTextBox1.Text.Trim() + ".xml"))
                {
                    errorProvider1.SetError(this.maskedTextBox1, " This name is already exist ");
                    return;
                }
                Variables.xDocument.Save(Variables.usersFolder + maskedTextBox1.Text.Trim() + ".xml");
                Variables.xmlFileName = Variables.usersFolder + maskedTextBox1.Text.Trim() + ".xml";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            this.Close();
        }
    }
}
