using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Xml.Linq;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Web;
using Leitner_Box;

namespace Leitner_Box
{
    public partial class Form1 : Form
    {
        #region Fields

        XElement selectedElement;
        bool EnableAutoComplete;

        #endregion

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public Form1()
        {
            InitializeComponent();
        }

        #region shortcut buttons

        void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.S && e.Control && buttonShowAnswer.Enabled)
            {
                buttonShowAnswer_Click(null, null);
            }
            else if (e.KeyCode == Keys.T && e.Control && buttonTrue.Enabled)
            {
                buttonTrue_Click(null, null);
            }
            else if (e.KeyCode == Keys.F && e.Control && buttonFalse.Enabled)
            {
                buttonFalse_Click(null, null);
            }
        }

        #endregion

        string ComputePersianDate(DateTime dateTime)
        {
            try
            {
                PersianCalendar persianCalendar = new PersianCalendar();
                var str = persianCalendar.GetYear(dateTime).ToString() + " / " +
                                            persianCalendar.GetMonth(dateTime).ToString() + " / " +
                                            persianCalendar.GetDayOfMonth(dateTime).ToString() + "   " +
                                            persianCalendar.GetHour(dateTime).ToString() + ":" +
                                            persianCalendar.GetMinute(dateTime).ToString() + ":" +
                                            persianCalendar.GetSecond(dateTime).ToString();
                return str;
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                return "";
            }
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region AutoComplete Events

        bool questionIsAvcitve = false;
        bool answerIsActive = false;

        private void textBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            EnableAutoComplete = true;
            TextBox textBox = sender as TextBox;

            if (textBox.Name == "textBoxNewQuestion")
            {
                questionIsAvcitve = true;
                answerIsActive = false;
            }
            else if (textBox.Name == "textBoxNewAnswer")
            {
                questionIsAvcitve = false;
                answerIsActive = true;
            }

            if ((int)e.KeyChar == 13 && listBoxAutoComplete.Visible)
            {
                try
                {
                    e.Handled = true;
                    textBox.Text = listBoxAutoComplete.Items[listBoxAutoComplete.SelectedIndex] as String;
                    listBoxAutoComplete.Visible = false;
                }
                catch { }
            }
        }

        private void textBoxKeyDown(object sender, KeyEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (e.Control && e.KeyCode == Keys.A)
            {
                textBox.SelectAll();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                listBoxAutoComplete.Visible = false;
            }
            else if (e.KeyCode == Keys.Down && listBoxAutoComplete.Visible)
            {
                try { listBoxAutoComplete.SelectedIndex++; }
                catch { }
            }
            else if (e.KeyCode == Keys.Up && listBoxAutoComplete.Visible)
            {
                try { listBoxAutoComplete.SelectedIndex--; }
                catch { }
            }
        }

        private void textBox_Enter(object sender, EventArgs e)
        {
            labelAddQuestionMessage.Text = "";
            labelAnswerToQuestionMessage.Text = "";
            labelSearchMessage.Text = "";
            (sender as TextBox).SelectAll();
            listBoxAutoComplete.Visible = false;
        }

        void listBoxAutoComplete_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (questionIsAvcitve)
                {
                    textBoxNewQuestion.Text = listBoxAutoComplete.Items[listBoxAutoComplete.SelectedIndex] as String;
                    textBoxNewQuestion.Focus();
                }
                else if (answerIsActive)
                {
                    textBoxNewAnswer.Text = listBoxAutoComplete.Items[listBoxAutoComplete.SelectedIndex] as String;
                    textBoxNewAnswer.Focus();
                }
            }
            catch { }
            finally
            {
                listBoxAutoComplete.Visible = false;
            }
        }

        #endregion

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region Messages

        void error(ref StackFrame file_info, string errorMassage)
        {
            try
            {
                if (file_info.GetFileName() == null)
                    MessageBox.Show(this, "Exception : " + errorMassage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show(this, "File : " + file_info.GetFileName() + "\nLine : " + file_info.GetFileLineNumber().ToString() + "\nException : " + errorMassage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch { }
        }

        void successful(string title, string message)
        {
            try
            {
                MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch { }
        }

        #endregion

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region Load XML file and apply Setting

        private void Form1_Shown(object sender, EventArgs e)
        {
            //Load the xml File
            try
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(Variables.usersFolder);
                if (directoryInfo.Exists)
                {
                    FileInfo[] users = directoryInfo.GetFiles("*.xml");
                    if (users.Count() <= 0)
                    {
                        newUser newUserForm = new newUser();
                        newUserForm.ShowDialog();
                        loadXML();
                    }
                    else if (users.Count() <= 1)
                    {
                        Variables.xmlFileName = users[0].FullName;

                        loadXML();

                        labelAddQuestionMessage.Text = "";
                    }
                    else if (users.Count() > 1)
                    {
                        selectUser selectUserForm = new selectUser();
                        selectUserForm.ShowDialog();
                        loadXML();
                    }
                }
                else
                {
                    newUser newUserForm = new newUser();
                    newUserForm.ShowDialog();
                    loadXML();
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        void addNodesToBox1()
        {
            treeView1.Nodes[0].Nodes.Clear();
            var questions = from q in Variables.xDocument.Descendants("Box")
                            where q.Attribute("ID").Value == "1"
                            from c in q.Descendants("Word")
                            select c;
            int i = 1;
            foreach (var question in questions)
            {
                treeView1.Nodes[0].Nodes.Add("Word" + question.Attribute("ID").Value, "Question " + i.ToString(), 3, 3);
                i++;
            }
        }

        void addNodesToBox2_3_4_5()
        {
            int j = 0;
            for (int i = 1; i < 5; i++)
            {
                //i+1 ==> box id
                //j ==> part numbers
                if (i + 1 == 2) j = 2;
                else if (i + 1 == 3) j = 5;
                else if (i + 1 == 4) j = 8;
                else if (i + 1 == 5) j = 14;

                for (int k = 1; k <= j; k++)
                {
                    treeView1.Nodes[i].Nodes[k - 1].Nodes.Clear();
                    var partQuestions = from q in Variables.xDocument.Descendants("Box")
                                        where q.Attribute("ID").Value == (i + 1).ToString()
                                        from c in q.Descendants("Part")
                                        where c.Attribute("ID").Value == k.ToString()
                                        from l in c.Descendants("Word")
                                        select l;
                    int g = 1;
                    foreach (var question in partQuestions)
                    {
                        treeView1.Nodes[i].Nodes[k - 1].Nodes.Add("Word" + question.Attribute("ID").Value, "Question " + g.ToString(), 3, 3);
                        g++;
                    }
                }
            }
        }

        void loadXML()
        {
            try
            {
                Variables.xDocument = XDocument.Load(Variables.xmlFileName);

                string userName = Variables.xmlFileName.Remove(Variables.xmlFileName.LastIndexOf('.'));
                userName = userName.Replace(Variables.usersFolder, "");
                userName = userName.Replace("\\", "");
                this.Text = Variables.title + userName;

                addNodesToBox1();
                addNodesToBox2_3_4_5();
                ApplySetting();

                treeView1.Enabled = tabControl1.Enabled = true;

                timer1.Start();
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                treeView1.Enabled = tabControl1.Enabled = false;
                timer1.Stop();
            }
        }

        void ApplySetting()
        {
            try
            {
                var Date = (from q in Variables.xDocument.Descendants("Date")
                            select q).First();
                if (Date.Value == "Christian")
                {
                    ToolStripMenuItemChristianDate.Checked = true;
                    ToolStripMenuItemPersianDate.Checked = false;
                }
                else
                {
                    ToolStripMenuItemChristianDate.Checked = false;
                    ToolStripMenuItemPersianDate.Checked = true;
                }

                var QuestionTextBoxes = (from q in Variables.xDocument.Descendants("QuestionTextBox")
                                         select q).First();
                if (QuestionTextBoxes.Value == "RightToLeft")
                    RightToLeftQuestion();
                else
                    LeftToRightQuestion();

                var AnswerTextBoxes = (from q in Variables.xDocument.Descendants("AnswerTextBox")
                                       select q).First();
                if (AnswerTextBoxes.Value == "RightToLeft")
                    RightToLeftAnswer();
                else
                    LeftToRightAnswer();
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        void RightToLeftQuestion()
        {
            ToolStripMenuItemLeftToRightQuestion.Checked = false;
            ToolStripMenuItemRightToLeftQuestion.Checked = true;

            textBoxNewQuestion.RightToLeft = RightToLeft.Yes;
            textBoxQuestion.RightToLeft = RightToLeft.Yes;
            textBoxSearchQuestion.RightToLeft = RightToLeft.Yes;
        }

        void LeftToRightQuestion()
        {
            ToolStripMenuItemLeftToRightQuestion.Checked = true;
            ToolStripMenuItemRightToLeftQuestion.Checked = false;

            textBoxNewQuestion.RightToLeft = RightToLeft.No;
            textBoxQuestion.RightToLeft = RightToLeft.No;
            textBoxSearchQuestion.RightToLeft = RightToLeft.No;
        }

        void RightToLeftAnswer()
        {
            ToolStripMenuItemRightToLeftAnswer.Checked = true;
            ToolStripMenuItemLeftToRightAnswer.Checked = false;

            textBoxNewAnswer.RightToLeft = RightToLeft.Yes;
            textBoxAnswer.RightToLeft = RightToLeft.Yes;
            textBoxSearchAnswer.RightToLeft = RightToLeft.Yes;
        }

        void LeftToRightAnswer()
        {
            ToolStripMenuItemRightToLeftAnswer.Checked = false;
            ToolStripMenuItemLeftToRightAnswer.Checked = true;

            textBoxNewAnswer.RightToLeft = RightToLeft.No;
            textBoxAnswer.RightToLeft = RightToLeft.No;
            textBoxSearchAnswer.RightToLeft = RightToLeft.No;
        }

        #endregion

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region XML Queries -> Updating XML file

        /// <summary>
        /// Adds an element to the xml file
        /// </summary>
        bool addToXMLAndTreeView(string boxID, string partID, string newQuestion, string newAnswer, string date, bool selectDestinationTreeNode)
        {
            try
            {
                XElement destinationBox;
                switch (boxID)
                {
                    case "1":
                        destinationBox = (from q in Variables.xDocument.Descendants("Box")
                                          where q.Attribute("ID").Value == "1"
                                          select q).First();
                        break;

                    case "2":
                    case "3":
                    case "4":
                    case "5":
                        destinationBox = (from q in Variables.xDocument.Descendants("Box")
                                          where q.Attribute("ID").Value == boxID
                                          from c in q.Descendants("Part")
                                          where c.Attribute("ID").Value == partID
                                          select c).First();
                        break;

                    case "DataBase":
                        destinationBox = (from q in Variables.xDocument.Descendants("DataBase")
                                          select q).First();
                        break;

                    default:
                        throw new Exception("Wrong box id");
                }

                int maxID = 0;
                try
                {
                    maxID = (from q in Variables.xDocument.Descendants("Word")
                             select (int)q.Attribute("ID")).Max();
                }
                catch { }
                maxID++;

                destinationBox.Add(
                        new XElement("Word", new XAttribute("ID", maxID),
                            new XAttribute("Question", newQuestion),
                            new XAttribute("Answer", newAnswer),
                            new XAttribute("Date", date))
                        );

                //Adds a new TreeNode to treeView
                TreeNode destinationTreeNode = new TreeNode();
                if (boxID == "1")
                {
                    destinationTreeNode = treeView1.Nodes.Find("Box1", false).First();
                    destinationTreeNode.Nodes.Add("Word" + maxID, "Question " + (destinationTreeNode.Nodes.Count + 1).ToString(), 3, 3);
                }
                else if (boxID != "DataBase")
                {
                    destinationTreeNode = treeView1.Nodes.Find("Box" + boxID + "Part" + partID, true).First();
                    destinationTreeNode.Nodes.Add("Word" + maxID, "Question " + (destinationTreeNode.Nodes.Count + 1).ToString(), 3, 3);
                }
                //\\
                try
                {
                    if (selectDestinationTreeNode && destinationTreeNode != null && boxID != "DataBase")
                    {
                        treeView1.CollapseAll();
                        treeView1.SelectedNode = treeView1.Nodes.Find("Word" + maxID, true).First();
                    }
                }
                catch { }

                Variables.xDocument.Save(Variables.xmlFileName);
                return true;
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Deletes an element from xml file
        /// </summary>
        bool deleteFromXMLAndTreeView(string wordID)
        {
            try
            {
                var destinationElement = (from q in Variables.xDocument.Descendants("Word")
                                          where q.Attribute("ID").Value == wordID
                                          select q).First();
                treeView1.Nodes.Find("Word" + wordID, true).First().Remove();
                destinationElement.Remove();
                Variables.xDocument.Save(Variables.xmlFileName);
                return true;
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
                return false;
            }
        }

        #endregion

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region tabControl1

        #region TabPages Enter Events

        private void tabPageInsertWord_Enter(object sender, EventArgs e)
        {
            listBoxAutoComplete.Visible = false;
            labelAddQuestionMessage.Text = "";
        }

        private void tabPageExplorer_Enter(object sender, EventArgs e)
        {
            labelAnswerToQuestionMessage.Text = "";
            EnableAutoComplete = false;
            listBoxAutoComplete.Visible = false;
        }

        private void tabPageStatistics_Enter(object sender, EventArgs e)
        {
            try
            {
                EnableAutoComplete = false;
                listBoxAutoComplete.Visible = false;

                labelStatisticsBox1.Text = (from q in Variables.xDocument.Descendants("Box")
                                            where q.Attribute("ID").Value == "1"
                                            from c in q.Descendants("Word")
                                            select c).Count().ToString();

                labelStatisticsBox2.Text = (from q in Variables.xDocument.Descendants("Box")
                                            where q.Attribute("ID").Value == "2"
                                            from c in q.Descendants("Word")
                                            select c).Count().ToString();

                labelStatisticsBox3.Text = (from q in Variables.xDocument.Descendants("Box")
                                            where q.Attribute("ID").Value == "3"
                                            from c in q.Descendants("Word")
                                            select c).Count().ToString();

                labelStatisticsBox4.Text = (from q in Variables.xDocument.Descendants("Box")
                                            where q.Attribute("ID").Value == "4"
                                            from c in q.Descendants("Word")
                                            select c).Count().ToString();

                labelStatisticsBox5.Text = (from q in Variables.xDocument.Descendants("Box")
                                            where q.Attribute("ID").Value == "5"
                                            from c in q.Descendants("Word")
                                            select c).Count().ToString();

                labelStatisticsDataBase.Text = (from q in Variables.xDocument.Descendants("DataBase")
                                                from c in q.Descendants("Word")
                                                select c).Count().ToString();

                labelAllWords.Text = (from q in Variables.xDocument.Descendants("Word")
                                      select q).Count().ToString();

                string dateString = (from q in Variables.xDocument.Descendants("StartTime") select q.Value).First();

                //Start date
                if (ToolStripMenuItemPersianDate.Checked)
                    labelStatisticsDate.Text = ComputePersianDate(DateTime.Parse(dateString));
                else
                    labelStatisticsDate.Text = dateString;
                //\\
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void tabPageSearch_Enter(object sender, EventArgs e)
        {
            EnableAutoComplete = false;
            listBoxAutoComplete.Visible = false;

            listViewSearch.Items.Clear();

            buttonSearchSave.Enabled = false;
            buttonSearchDelete.Enabled = false;
            buttonSearchMoveToBox1.Enabled = false;
            textBoxSearchAnswer.Enabled = false;
            textBoxSearchQuestion.Enabled = false;

            textBoxSearchAnswer.Text = textBoxSearchQuestion.Text = "";
            labelSearchMessage.Text = "";
        }

        #endregion

        #region Buttons

        private void buttonAddNewWord_Click(object sender, EventArgs e)
        {
            try
            {
                labelAddQuestionMessage.Text = "";

                ////////////////////////////////////
                errorProvider1.Clear();

                labelAddQuestionMessage.Text = "";

                if (textBoxNewQuestion.Text.Trim() == "")
                {
                    errorProvider1.SetError(textBoxNewQuestion, "Please fill in this textbox");
                    return;
                }
                else if (textBoxNewAnswer.Text.Trim() == "")
                {
                    errorProvider1.SetError(textBoxNewAnswer, "Please fill in this textbox");
                    return;
                }

                var exist = (from q in Variables.xDocument.Descendants("Word")
                             where q.Attribute("Question").Value.ToLower() == textBoxNewQuestion.Text.Trim().ToLower()
                             select q).Count();

                if (exist > 0)
                {
                    errorProvider1.SetError(this.textBoxNewQuestion, "There is a same question in the Leitner Box");
                    return;
                }
                ////////////////////////////////////

                string boxID = labelBoxID.Text;
                string partID = "";
                try
                {
                    partID = labelPartID.Text;
                }
                catch { }

                if (!addToXMLAndTreeView(boxID, partID, textBoxNewQuestion.Text.Trim(), textBoxNewAnswer.Text.Trim(), DateTime.Now.ToString().Replace("ب.ظ", "PM").Replace("ق.ظ", "AM"), true))
                {
                    labelAddQuestionMessage.Text = "Error in adding data to XML file";
                    return;
                }

                if (partID != "")
                    labelAddQuestionMessage.Text = "The question added to Box " + boxID + " -> " + "Part " + partID + " successfully";
                else if (boxID == "DataBase")
                    labelAddQuestionMessage.Text = "The question added to " + boxID + " successfully";
                else
                    labelAddQuestionMessage.Text = "The question added to Box " + boxID + " successfully";
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void buttonDelete1_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Are you sure ?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

                string question = this.selectedElement.Attribute("Question").Value;
                string id = this.selectedElement.Attribute("ID").Value;

                this.selectedElement.Remove();

                Variables.xDocument.Save(Variables.xmlFileName);

                treeView1.Nodes.Find("Word" + id, true).First().Remove();

                labelAnswerToQuestionMessage.Text = "The question deleted successfully";
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void buttonSave1_Click(object sender, EventArgs e)
        {
            try
            {
                labelAnswerToQuestionMessage.Text = "";

                ////////////////////////////////////
                errorProvider1.Clear();

                if (textBoxQuestion.Text.Trim() == "")
                {
                    errorProvider1.SetError(textBoxQuestion, "Please fill this textbox out");
                    return;
                }
                else if (textBoxAnswer.Text.Trim() == "")
                {
                    errorProvider1.SetError(textBoxAnswer, "Please fill this textbox out");
                    return;
                }
                ////////////////////////////////////

                this.selectedElement.Attribute("Question").Value = textBoxQuestion.Text.Trim();
                this.selectedElement.Attribute("Answer").Value = textBoxAnswer.Text.Trim();

                Variables.xDocument.Save(Variables.xmlFileName);

                labelAnswerToQuestionMessage.Text = "The changes saved successfully";
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void buttonFalse_Click(object sender, EventArgs e)
        {
            try
            {
                labelAnswerToQuestionMessage.Text = "";
                if (labelBoxID.Text == "1") return;

                string question = this.selectedElement.Attribute("Question").Value;
                string answer = this.selectedElement.Attribute("Answer").Value;
                string date = this.selectedElement.Attribute("Date").Value;
                string wordID = treeView1.SelectedNode.Name.Replace("Word", "").Trim();

                var boxID = "1";

                try
                {
                    boxID = this.selectedElement.Parent.Parent.Attribute("ID").Value;
                }
                catch { }
                var partID = "";
                try
                {
                    partID = this.selectedElement.Parent.Attribute("ID").Value;
                }
                catch { }
                deleteFromXMLAndTreeView(wordID);
                addToXMLAndTreeView("1", "", question, answer, date, false);

                labelAnswerToQuestionMessage.Text = "The question moved to Box1";
            }
            catch (Exception ex)
            {
                labelAnswerToQuestionMessage.Text = "";
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void buttonTrue_Click(object sender, EventArgs e)
        {
            try
            {
                labelAnswerToQuestionMessage.Text = "";
                if (labelPartID.Text != "" && labelPartID.Text != "1") return;

                string question = this.selectedElement.Attribute("Question").Value;
                string answer = this.selectedElement.Attribute("Answer").Value;
                string date = this.selectedElement.Attribute("Date").Value;

                string boxID = labelBoxID.Text;
                string wordID = treeView1.SelectedNode.Name.Replace("Word", "").Trim();

                switch (labelBoxID.Text)
                {
                    ////////////////////////////////
                    //Moves the word to Box2 Part2
                    case "1":
                        boxID = "1";
                        addToXMLAndTreeView("2", "2", question, answer, date, false);
                        deleteFromXMLAndTreeView(wordID);
                        labelAnswerToQuestionMessage.Text = "The question moved to Bax2 Part2";
                        break;

                    ////////////////////////////////
                    //Moves the word to Box3 Part5
                    case "2":
                        addToXMLAndTreeView("3", "5", question, answer, date, false);
                        deleteFromXMLAndTreeView(wordID);
                        labelAnswerToQuestionMessage.Text = "The question moved to Bax3 Part5";
                        break;

                    ////////////////////////////////
                    case "3":
                        addToXMLAndTreeView("4", "8", question, answer, date, false);
                        deleteFromXMLAndTreeView(wordID);
                        labelAnswerToQuestionMessage.Text = "The question moved to Bax4 Part8";
                        break;

                    ////////////////////////////////
                    case "4":
                        addToXMLAndTreeView("5", "14", question, answer, date, false);
                        deleteFromXMLAndTreeView(wordID);
                        labelAnswerToQuestionMessage.Text = "The question moved to Bax5 Part14";
                        break;

                    ////////////////////////////////
                    case "5":
                        addToXMLAndTreeView("DataBase", "", question, answer, date, false);
                        deleteFromXMLAndTreeView(wordID);
                        labelAnswerToQuestionMessage.Text = "The question moved to Data Base";
                        break;

                    default:
                        throw new Exception("wrong box id");
                }
                Variables.xDocument.Save(Variables.xmlFileName);
            }
            catch (Exception ex)
            {
                labelAnswerToQuestionMessage.Text = "Error";
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void buttonSearchSave_Click(object sender, EventArgs e)
        {
            try
            {
                labelSearchMessage.Text = "";

                string id = listViewSearch.SelectedItems[0].Name.Replace("Item", "");

                var element = (from q in Variables.xDocument.Descendants("Word")
                               where q.Attribute("ID").Value == id
                               select q).First();

                element.Attribute("Question").Value = textBoxSearchQuestion.Text.Trim();
                element.Attribute("Answer").Value = textBoxSearchAnswer.Text.Trim();

                Variables.xDocument.Save(Variables.xmlFileName);
                labelSearchMessage.Text = "The changes saved successfully .";
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void buttonSearchMoveToBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Are you sure ?", "Move", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

                string id = listViewSearch.SelectedItems[0].Name.Replace("Item", "");

                var element = (from q in Variables.xDocument.Descendants("Word")
                               where q.Attribute("ID").Value == id
                               select q).First();
                try
                {
                    treeView1.Nodes.Find("Word" + id, true).First().Remove();
                }
                catch { }
                try
                {
                    addToXMLAndTreeView("1", "", element.Attribute("Question").Value, element.Attribute("Answer").Value, element.Attribute("Date").Value, true);
                }
                catch { }
                element.Remove();

                Variables.xDocument.Save(Variables.xmlFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void buttonSearchDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Are you sure ?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

                string id = listViewSearch.SelectedItems[0].Name.Replace("Item", "");

                var element = (from q in Variables.xDocument.Descendants("Word")
                               where q.Attribute("ID").Value == id
                               select q).First();

                element.Remove();

                try
                {
                    listViewSearch.Items.Find("Item" + id, true).First().Remove();
                }
                catch { }
                try
                {
                    treeView1.Nodes.Find("Word" + id, true).First().Remove();
                }
                catch { }

                Variables.xDocument.Save(Variables.xmlFileName);
                labelSearchMessage.Text = "The question has been deleted successfully .";
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void buttonShowAnswer_Click(object sender, EventArgs e)
        {
            try
            {
                textBoxAnswer.Enabled = true;
                textBoxAnswer.Text = this.selectedElement.Attribute("Answer").Value.Replace("\n", Environment.NewLine);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        #endregion

        #region Search Methods

        private void textBoxSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                listViewSearch.Items.Clear();
                if (textBoxSearch.Text.Trim().Equals("") || textBoxSearch.Text.Trim().Length < 2) return;

                if (checkBoxSearchInQuestion.Checked && checkBoxSearchInAnswer.Checked)
                {
                    var words = from q in Variables.xDocument.Descendants("Word")
                                where q.Attribute("Question").Value.ToLower().Contains(textBoxSearch.Text.Trim().ToLower()) || q.Attribute("Answer").Value.ToLower().Contains(textBoxSearch.Text.Trim().ToLower())
                                select q;

                    if (words.Count() <= 0)
                    {
                        buttonSearchSave.Enabled = false;
                        buttonSearchDelete.Enabled = false;
                        buttonSearchMoveToBox1.Enabled = false;
                        textBoxSearchAnswer.Enabled = false;
                        textBoxSearchQuestion.Enabled = false;

                        textBoxSearchAnswer.Text = textBoxSearchQuestion.Text = "";
                        return;
                    }

                    foreach (var word in words)
                    {
                        string date = word.Attribute("Date").Value;
                        if (ToolStripMenuItemPersianDate.Checked)
                            date = ComputePersianDate(DateTime.Parse(date));

                        ListViewItem item = new ListViewItem(new string[] 
                        { word.Attribute("Question").Value, 
                          word.Attribute("Answer").Value, 
                          date });
                        item.Name = "Item" + word.Attribute("ID").Value;
                        listViewSearch.Items.Add(item);
                    }
                }
                else if (checkBoxSearchInQuestion.Checked && !checkBoxSearchInAnswer.Checked)
                {
                    var words = from q in Variables.xDocument.Descendants("Word")
                                where q.Attribute("Question").Value.ToLower().Contains(textBoxSearch.Text.Trim().ToLower())
                                select q;

                    if (words.Count() <= 0)
                    {
                        buttonSearchSave.Enabled = false;
                        buttonSearchDelete.Enabled = false;
                        buttonSearchMoveToBox1.Enabled = false;
                        textBoxSearchAnswer.Enabled = false;
                        textBoxSearchQuestion.Enabled = false;
                        textBoxSearchAnswer.Text = textBoxSearchQuestion.Text = "";
                        return;
                    }

                    foreach (var word in words)
                    {
                        string date = word.Attribute("Date").Value;
                        if (ToolStripMenuItemPersianDate.Checked)
                            date = ComputePersianDate(DateTime.Parse(date));

                        ListViewItem item = new ListViewItem(new string[] 
                        { word.Attribute("Question").Value, 
                          word.Attribute("Answer").Value, 
                          date });
                        item.Name = "Item" + word.Attribute("ID").Value;
                        listViewSearch.Items.Add(item);
                    }
                }
                else if (!checkBoxSearchInQuestion.Checked && checkBoxSearchInAnswer.Checked)
                {
                    var words = from q in Variables.xDocument.Descendants("Word")
                                where q.Attribute("Answer").Value.ToLower().Contains(textBoxSearch.Text.Trim().ToLower())
                                select q;

                    if (words.Count() <= 0)
                    {
                        buttonSearchSave.Enabled = false;
                        buttonSearchDelete.Enabled = false;
                        buttonSearchMoveToBox1.Enabled = false;
                        textBoxSearchAnswer.Enabled = false;
                        textBoxSearchQuestion.Enabled = false;
                        textBoxSearchAnswer.Text = textBoxSearchQuestion.Text = "";
                        return;
                    }

                    foreach (var word in words)
                    {
                        string date = word.Attribute("Date").Value;
                        if (ToolStripMenuItemPersianDate.Checked)
                            date = ComputePersianDate(DateTime.Parse(date));

                        ListViewItem item = new ListViewItem(new string[] 
                        { word.Attribute("Question").Value, 
                          word.Attribute("Answer").Value, 
                          date });
                        item.Name = "Item" + word.Attribute("ID").Value;
                        listViewSearch.Items.Add(item);
                    }
                }
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void listViewSearch_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {
                treeView1.CollapseAll();

                string id = ((ListView)sender).SelectedItems[0].Name.Replace("Item", "");

                try
                {
                    this.treeView1.SelectedNode =
                        this.treeView1.Nodes.Find("Word" + id, true).First();
                }
                catch
                {
                    this.treeView1.SelectedNode = this.treeView1.Nodes.Find("DataBase", true).First();
                }

                buttonSearchSave.Enabled = true;
                buttonSearchDelete.Enabled = true;
                buttonSearchMoveToBox1.Enabled = true;
                textBoxSearchAnswer.Enabled = true;
                textBoxSearchQuestion.Enabled = true;

                textBoxSearchQuestion.Text = Variables.xDocument.Descendants("Word").Where(q => (q.Attribute("ID").Value == id)).First().Attribute("Question").Value;
                textBoxSearchAnswer.Text = Variables.xDocument.Descendants("Word").Where(q => (q.Attribute("ID").Value == id)).First().Attribute("Answer").Value;
            }
            catch
            {
                buttonSearchSave.Enabled = false;
                buttonSearchDelete.Enabled = false;
                buttonSearchMoveToBox1.Enabled = false;
                textBoxSearchAnswer.Enabled = false;
                textBoxSearchQuestion.Enabled = false;
            }
        }

        private void textBoxNewQuestion_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (!EnableAutoComplete)
                {
                    listBoxAutoComplete.Visible = false;
                    return;
                }

                if (textBoxNewQuestion.Text.Trim() == "" || textBoxNewQuestion.Text.Trim().Length <= 1)
                {
                    listBoxAutoComplete.Visible = false;
                    return;
                }

                var elements = from q in Variables.xDocument.Descendants("Word")
                               where q.Attribute("Question").Value.ToLower().Contains(textBoxNewQuestion.Text.ToLower().Trim())
                                || q.Attribute("Answer").Value.ToLower().Contains(textBoxNewQuestion.Text.ToLower().Trim())
                               select q;

                if (elements.Count() == 0)
                {
                    listBoxAutoComplete.Visible = false;
                    return;
                }

                listBoxAutoComplete.Items.Clear();
                listBoxAutoComplete.Location = new Point(107, 294);//107, 294
                listBoxAutoComplete.Visible = true;
                listBoxAutoComplete.RightToLeft = textBoxNewQuestion.RightToLeft;

                foreach (var item in elements)
                {
                    listBoxAutoComplete.Items.Add(item.Attribute("Question").Value);
                }
            }
            catch { listBoxAutoComplete.Visible = false; }
        }

        private void textBoxNewAnswer_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (!EnableAutoComplete)
                {
                    listBoxAutoComplete.Visible = false;
                    return;
                }

                if (textBoxNewAnswer.Text.Trim() == "" || textBoxNewAnswer.Text.Trim().Length <= 1)
                {
                    listBoxAutoComplete.Visible = false;
                    return;
                }

                var elements = from q in Variables.xDocument.Descendants("Word")
                               where q.Attribute("Answer").Value.ToLower().Contains(textBoxNewAnswer.Text.ToLower().Trim())
                                || q.Attribute("Question").Value.ToLower().Contains(textBoxNewAnswer.Text.ToLower().Trim())
                               select q;

                if (elements.Count() == 0)
                {
                    listBoxAutoComplete.Visible = false;
                    return;
                }

                listBoxAutoComplete.Items.Clear();
                listBoxAutoComplete.Location = new Point(107, 154);//99, 162
                listBoxAutoComplete.Visible = true;
                listBoxAutoComplete.RightToLeft = textBoxNewAnswer.RightToLeft;

                foreach (var item in elements)
                {
                    listBoxAutoComplete.Items.Add(item.Attribute("Answer").Value);
                }
            }
            catch { }
        }

        #endregion

        #endregion

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region TreeView

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                string boxID = "1";
                string partID = "";

                labelAddQuestionMessage.Text = labelAnswerToQuestionMessage.Text = "";
                listBoxAutoComplete.Visible = false;
                EnableAutoComplete = false;

                //بدست آوردن تعداد سوال در گره انتخاب شده
                #region Obtaining the numbers of selected node's descendants

                if (e.Node.Name.Contains("Word"))
                {
                    try
                    {
                        label2WordsCount.Text = "1";
                        if (e.Node.Parent.Name.Contains("Part"))
                        {
                            partID = Regex.Replace(e.Node.Parent.Name, @"Box.Part", "", RegexOptions.IgnoreCase);
                            boxID = Regex.Replace(e.Node.Parent.Parent.Name, @"Box", "", RegexOptions.IgnoreCase);
                        }
                        else
                        {
                            partID = "";
                            boxID = "1";
                        }
                    }
                    catch
                    {
                        label2WordsCount.Text = "";
                    }
                }
                else if (e.Node.Name.Contains("Part"))
                {
                    try
                    {
                        boxID = Regex.Replace(e.Node.Parent.Name, @"Box", "", RegexOptions.IgnoreCase);
                        partID = Regex.Replace(e.Node.Name, @"Box.Part", "", RegexOptions.IgnoreCase);
                        label2WordsCount.Text = (from q in Variables.xDocument.Descendants("Box")
                                                 where q.Attribute("ID").Value == boxID
                                                 from c in q.Descendants("Part")
                                                 where c.Attribute("ID").Value == partID
                                                 from l in c.Descendants("Word")
                                                 select l).Count().ToString();
                    }
                    catch
                    {
                        label2WordsCount.Text = "";
                    }
                }
                else if (e.Node.Name.Contains("Box"))
                {
                    try
                    {
                        boxID = Regex.Replace(e.Node.Name, @"Box", "", RegexOptions.IgnoreCase);
                        if (boxID == "1") partID = "";
                        else partID = "1";
                        label2WordsCount.Text = (from q in Variables.xDocument.Descendants("Box")
                                                 where q.Attribute("ID").Value == boxID
                                                 from l in q.Descendants("Word")
                                                 select l).Count().ToString();
                    }
                    catch
                    {
                        label2WordsCount.Text = "";
                    }
                }
                else if (e.Node.Name.Contains("DataBase"))
                {
                    try
                    {
                        boxID = e.Node.Name;
                        partID = "";
                        label2WordsCount.Text = (from q in Variables.xDocument.Descendants("DataBase")
                                                 from l in q.Descendants("Word")
                                                 select l).Count().ToString();
                    }
                    catch
                    {
                        label2WordsCount.Text = "";
                    }
                }

                labelBoxID.Text = boxID;
                labelPartID.Text = partID;

                #endregion

                #region Getting the question and answer of selected node and showing it in the textboxes

                labelRegDate.Text = "00 / 00 / 00";
                textBoxQuestion.Enabled = false;
                textBoxAnswer.Text = "";
                textBoxAnswer.Enabled = false;
                buttonTrue.Enabled = false;
                buttonFalse.Enabled = false;
                buttonSave1.Enabled = false;
                buttonShowAnswer.Enabled = false;
                buttonDelete1.Enabled = false;
                textBoxNewQuestion.Text = "";
                textBoxNewAnswer.Text = "";
                try
                {
                    string id = e.Node.Name.Replace("Word", "");
                    this.selectedElement = (from q in Variables.xDocument.Descendants("Word")
                                            where q.Attribute("ID").Value == id
                                            select q).First();

                    textBoxQuestion.Text = selectedElement.Attribute("Question").Value;

                    textBoxNewQuestion.Text = textBoxQuestion.Text;
                    textBoxNewAnswer.Text = selectedElement.Attribute("Answer").Value;

                    textBoxQuestion.Enabled = true;
                    buttonSave1.Enabled = true;
                    buttonShowAnswer.Enabled = true;
                    buttonDelete1.Enabled = true;

                    if (partID == "1" || partID == "")
                    {
                        buttonTrue.Enabled = true;
                        buttonFalse.Enabled = true;
                    }

                    //Register date
                    if (ToolStripMenuItemPersianDate.Checked)
                        labelRegDate.Text = ComputePersianDate(DateTime.Parse(this.selectedElement.Attribute("Date").Value));
                    else
                        labelRegDate.Text = this.selectedElement.Attribute("Date").Value;
                    //\\
                }
                catch
                {
                    textBoxQuestion.Text = "";
                }

                #endregion

            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void treeView1_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                EnableAutoComplete = false;
                listBoxAutoComplete.Visible = false;

                treeView1.SelectedNode = treeView1.GetNodeAt(new Point(e.X, e.Y));
                if (e.Button == MouseButtons.Right)
                    contextMenuStrip1.Show(this.treeView1, new Point(e.X, e.Y));

                ToolStripMenuItemShiftUp.Enabled = true;
                ToolStripMenuItemExport.Enabled = true;
                ToolStripMenuItemDelete.Enabled = false;
                ToolStripMenuItemMoveTo.Enabled = false;

                if (treeView1.SelectedNode.Name.Contains("Word"))
                {
                    ToolStripMenuItemDelete.Enabled = true;
                    ToolStripMenuItemExport.Enabled = false;
                    ToolStripMenuItemShiftUp.Enabled = false;
                    ToolStripMenuItemMoveTo.Enabled = true;
                }
                else if (treeView1.SelectedNode.Name.Contains("Part"))
                {
                    ToolStripMenuItemDelete.Enabled = false;
                    ToolStripMenuItemExport.Enabled = false;
                    ToolStripMenuItemShiftUp.Enabled = false;
                }

                if (labelBoxID.Text == "1" || labelBoxID.Text == "DataBase")
                    ToolStripMenuItemShiftUp.Enabled = false;
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        #endregion

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region MenuStrip

        private void optimizeXmlFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                var elements = from q in Variables.xDocument.Descendants("Word")
                               select q;
                int counter = 0;
                foreach (var element in elements)
                {
                    counter++;
                    element.SetAttributeValue("ID", counter);
                }
                Variables.xDocument.Save(Variables.xmlFileName);
                loadXML();
                successful("Optimization", "The XML file has been optimized successfully");
                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ToolStripMenuItemRightToLeftQuestion_Click(object sender, EventArgs e)
        {
            try
            {
                RightToLeftQuestion();

                var QuestionTextBoxes = (from q in Variables.xDocument.Descendants("QuestionTextBox")
                                         select q).First();
                QuestionTextBoxes.Value = "RightToLeft";
                Variables.xDocument.Save(Variables.xmlFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void ToolStripMenuItemLeftToRightQuestion_Click(object sender, EventArgs e)
        {
            try
            {
                LeftToRightQuestion();

                var QuestionTextBoxes = (from q in Variables.xDocument.Descendants("QuestionTextBox")
                                         select q).First();
                QuestionTextBoxes.Value = "";
                QuestionTextBoxes.Value = "LeftToRight";
                Variables.xDocument.Save(Variables.xmlFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void ToolStripMenuItemRightToLeftAnswer_Click(object sender, EventArgs e)
        {
            try
            {
                RightToLeftAnswer();

                var AnswerTextBoxes = (from q in Variables.xDocument.Descendants("AnswerTextBox")
                                       select q).First();
                AnswerTextBoxes.Value = "RightToLeft";
                Variables.xDocument.Save(Variables.xmlFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void ToolStripMenuItemLeftToRightAnswer_Click(object sender, EventArgs e)
        {
            try
            {
                LeftToRightAnswer();

                var AnswerTextBoxes = (from q in Variables.xDocument.Descendants("AnswerTextBox")
                                       select q).First();
                AnswerTextBoxes.Value = "LeftToRight";
                Variables.xDocument.Save(Variables.xmlFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void ToolStripMenuItemChristianDate_Click(object sender, EventArgs e)
        {
            try
            {
                ToolStripMenuItemChristianDate.Checked = true;
                ToolStripMenuItemPersianDate.Checked = false;

                var Date = (from q in Variables.xDocument.Descendants("Date")
                            select q).First();
                Date.Value = "Christian";
                Variables.xDocument.Save(Variables.xmlFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void ToolStripMenuItemPersianDate_Click(object sender, EventArgs e)
        {
            try
            {
                ToolStripMenuItemChristianDate.Checked = false;
                ToolStripMenuItemPersianDate.Checked = true;

                var Date = (from q in Variables.xDocument.Descendants("Date")
                            select q).First();
                Date.Value = "Persian";
                Variables.xDocument.Save(Variables.xmlFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void exportAllWordsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;
                this.Cursor = Cursors.WaitCursor;
                IEnumerable<XElement> elements;
                elements = from q in Variables.xDocument.Descendants("Word") select q;
                exportWords(ref elements, saveFileDialog1.FileName);
                this.Cursor = Cursors.Default;
                successful("File Saving", "The file has been saved successfully");
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void ToolStripMenuItemAboutMe_Click(object sender, EventArgs e)
        {
            aboutMe aboutMeForm = new aboutMe();
            aboutMeForm.ShowDialog();
        }

        private void ToolStripMenuItemCreateNewUser_Click(object sender, EventArgs e)
        {
            newUser newUserForm = new newUser();
            newUserForm.ShowDialog();
            loadXML();
        }

        private void ToolStripMenuItemSelectUser_Click(object sender, EventArgs e)
        {
            try
            {
                selectUser selectUserForm = new selectUser();
                selectUserForm.ShowDialog();
                loadXML();
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void ToolStripMenuItemMoveTo_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem toolStripMenuItem = sender as ToolStripMenuItem;

            string question = this.selectedElement.Attribute("Question").Value;
            string answer = this.selectedElement.Attribute("Answer").Value;
            string date = this.selectedElement.Attribute("Date").Value;
            string id = this.selectedElement.Attribute("ID").Value;

            //ToolStripMenuItemBox2Part1
            string boxID = Regex.Replace(toolStripMenuItem.Name, @"(ToolStripMenuItem)|(Box)|(Part..*)", "", RegexOptions.IgnoreCase);
            string partID = Regex.Replace(toolStripMenuItem.Name, @"(ToolStripMenuItem)|(Box.)|(Part)", "", RegexOptions.IgnoreCase);

            this.selectedElement.Remove();
            treeView1.Nodes.Find("Word" + id, true).First().Remove();

            addToXMLAndTreeView(boxID, partID, question, answer, date, true);

            if (boxID == "1")
                labelAddQuestionMessage.Text = "The word moved to Box 1 successfully";
            else
                labelAddQuestionMessage.Text = "The word moved to Box " + boxID.ToString() + " Part " + partID.ToString() + " successfully";
        }


        #endregion

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region ContextMenuStrip

        private void toolStripMenuItemDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Are you sure ?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                this.selectedElement.Remove();
                treeView1.SelectedNode.Remove();
                Variables.xDocument.Save(Variables.xmlFileName);
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        private void toolStripMenuItemShiftUp_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Are you sure ?", "Shift up", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

                IEnumerable<XElement> box = from q in Variables.xDocument.Descendants("Box")
                                            where q.Attribute("ID").Value == labelBoxID.Text
                                            select q;

                XElement part1 = (from q in box.Descendants("Part")
                                  where q.Attribute("ID").Value == "1"
                                  select q).First();

                XElement part2 = (from q in box.Descendants("Part")
                                  where q.Attribute("ID").Value == "2"
                                  select q).First();

                part1.Add(part2.Descendants("Word"));

                for (int i = 2; i <= box.Descendants("Part").Count(); i++)
                {
                    try
                    {
                        var part = box.Descendants("Part").Where(q => (q.Attribute("ID").Value == i.ToString())).First();
                        part.Elements("Word").Remove();
                        part.Add(box.Descendants("Part").Where(q => (q.Attribute("ID").Value == (i + 1).ToString())).First().Elements("Word"));
                    }
                    catch { continue; }
                }
                Variables.xDocument.Save(Variables.xmlFileName);

                Form1_Shown(null, null);
                successful("Shift", "Leitner Box has been updated");
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        void toolStripMenuItemExport_Click(object sender, EventArgs e)
        {
            try
            {
                IEnumerable<XElement> elements;
                if (labelBoxID.Text == "DataBase")
                {
                    elements = from q in Variables.xDocument.Descendants("DataBase")
                               from c in q.Descendants("Word")
                               select c;
                }
                else
                {
                    elements = from q in Variables.xDocument.Descendants("Box")
                               where q.Attribute("ID").Value == labelBoxID.Text
                               from c in q.Descendants("Word")
                               select c;
                }

                if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;

                exportWords(ref elements, saveFileDialog1.FileName);

                successful("File Saving", "The file has been saved successfully");
            }
            catch (Exception ex)
            {
                StackFrame file_info = new StackFrame(true);
                error(ref file_info, ex.Message);
            }
        }

        void exportWords(ref IEnumerable<XElement> elements, string filename)
        {
            using (StreamWriter sw = new StreamWriter(filename))
            {
                FileInfo fileInfo = new FileInfo(Variables.xmlFileName);

                string date;
                if (ToolStripMenuItemChristianDate.Checked)
                    date = DateTime.Now.ToString();
                else
                    date = ComputePersianDate(DateTime.Now);

                sw.WriteLine("<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">");
                sw.WriteLine("<html xmlns=\"http://www.w3.org/1999/xhtml\">");
                sw.WriteLine("<head>");
                sw.WriteLine("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />");
                sw.WriteLine("<title>Leitner Box words --&gt;"
                        + fileInfo.Name.Remove(fileInfo.Name.LastIndexOf('.'))
                        + "  \n"
                        + date + " </title>");
                sw.WriteLine("</head>");
                sw.WriteLine("<!-- Created by Mohammad Dayyan http://www.mds-soft.persianblog.ir/ -->");
                sw.WriteLine("<body bgcolor=\"#F4F3EE\">");
                sw.WriteLine("<table width=\"100%\" border=\"1\" align=\"center\" cellpadding=\"3\" cellspacing=\"1\">  <tr><th>No.</th><th>Question</th><th>Answer</th></tr>");

                int i = 0;
                foreach (var word in elements)
                {
                    i++;
                    sw.WriteLine("\t<tr>");
                    sw.WriteLine("\t\t<td align=\"center\">" + i.ToString() + "</td>");
                    if (ToolStripMenuItemRightToLeftQuestion.Checked)
                        sw.WriteLine("\t\t<td align=\"center\" dir=\"rtl\">" + HttpUtility.HtmlDecode(word.Attribute("Question").Value).Replace("\n", "<br />") + "</td>");
                    else
                        sw.WriteLine("\t\t<td align=\"center\">" + HttpUtility.HtmlDecode(word.Attribute("Question").Value).Replace("\n", "<br />") + "</td>");
                    if (ToolStripMenuItemRightToLeftAnswer.Checked)
                        sw.WriteLine("\t\t<td align=\"center\" dir=\"rtl\">" + HttpUtility.HtmlDecode(word.Attribute("Answer").Value).Replace("\n", "<br />") + "</td>");
                    else
                        sw.WriteLine("\t\t<td align=\"center\">" + HttpUtility.HtmlDecode(word.Attribute("Answer").Value).Replace("\n", "<br />") + "</td>");
                    sw.WriteLine("\t</tr>");
                }

                sw.WriteLine("\n</table>\n</body>\n</html>");
                sw.Close();
            }
        }

        #endregion

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region Timer

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (ToolStripMenuItemPersianDate.Checked)
                labelNewWordDate.Text = ComputePersianDate(DateTime.Now);
            else
                labelNewWordDate.Text = DateTime.Now.ToString();
        }

        #endregion

        private void textBoxNewQuestion_Or_NewAnswer_Leave(object sender, EventArgs e)
        {
            listBoxAutoComplete.Visible = false;
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    }
}
