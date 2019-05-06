using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;

namespace Leitner_Box
{
    public static class Variables
    {
        public static XDocument xDocument;
        public static string xmlFileName = "";
        public static string usersFolder = Application.StartupPath + "\\users\\";
        public readonly static string title = "Leitner Box --> ";
    }
}
