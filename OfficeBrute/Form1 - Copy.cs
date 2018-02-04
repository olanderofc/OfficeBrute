using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.IO;

namespace OfficeBrute
{
    delegate void SetTextCallback(string text);

    public partial class Form1 : Form
    {
        string m_fileName = "";
        string m_wordlistFile = "";

        UInt64 m_wordlistRows = 0;
        UInt64 THREADS = 4;
        UInt64 m_first = 0;
        UInt64 m_second = 0;
        UInt64 m_third = 0;
        UInt64 m_fourth = 0;
        
        bool cracked = false;

        public Form1()
        {
            InitializeComponent();
        }
        private void SetText(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.lblCurrentPass.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.lblCurrentPass.Text = text;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                //btnCrackPasswd.Enabled = true;
            }
        }
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            if (textBox1.Text != null)
            {
                m_fileName = textBox1.Text;
                m_wordlistRows = (UInt64) File.ReadLines(m_wordlistFile).Count();

                start();
            }
        }

        public void start()
        {

            //Rows is 23403 / 4 =  5850
            UInt64 index = m_wordlistRows / THREADS;

             m_first = 0;
             m_second = m_first + index; // 5850
             m_third = m_second + index; // 5850 + 5850
             m_fourth = m_third + index; // 5850 + 5850 + 5850
            while (m_fourth < m_wordlistRows)
            {
               m_fourth--;
            }
 
            Thread t = new Thread(new ParameterizedThreadStart(bruteforce));
            UInt64 end =  m_first + index - 1;
            UInt64[] first_data = {m_first, end};
            t.Start((object)first_data);

            Thread u = new Thread(new ParameterizedThreadStart(bruteforce));
            end = m_second + index - 1;
            UInt64[] second_data = { m_second, end };
            u.Start((object)second_data);

            Thread r = new Thread(new ParameterizedThreadStart(bruteforce));
            end = m_third + index - 1;
            UInt64[] third_data = { m_third, end };
            r.Start((object)third_data);

            Thread s = new Thread(new ParameterizedThreadStart(bruteforce));
            end = m_fourth + index - 1;
            UInt64[] fourth_data = { m_fourth, end };
            s.Start((object)fourth_data);

        }

        public void bruteforce(object fileReadParams)
        {

            UInt64[] parameters = (UInt64[])fileReadParams;
            string fileName = m_fileName;
            string wordlistFile = m_wordlistFile;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            object nullobj = System.Reflection.Missing.Value;

            /*
             * 
             */
            string line;
            System.IO.StreamReader wordlist = new System.IO.StreamReader(wordlistFile);

            object password = "tmp";
            object FName = fileName;
            Microsoft.Office.Interop.Word.Document aDoc = null;
            wordApp.Visible = false;
            aDoc = wordApp.Documents.Open(ref FName, ref nullobj, ref nullobj, ref nullobj, ref password, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj);
            aDoc.Activate();

            if (aDoc.ProtectionType != Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
            {
                MessageBox.Show("Document is protected. Will now try to crack it");
            }

            string finalPassword = "";
            bool cracked = false;

            UInt64 counter = 0;
            UInt64 start = parameters[0];
            UInt64 end = parameters[1];
            //for (int i = 0; i < passwords.Length; i++)
            while ((line = wordlist.ReadLine()) != null)
            {

                try
                {
                    this.SetText(line);
                    object passwd = line;
                    counter++;
                    aDoc.Unprotect(ref passwd);
                    MessageBox.Show("Success. Document cracked");
                    finalPassword = line;
                    cracked = true;
                    break;
                }
                catch (Exception e)
                {
                    //MessageBox.Show("Wrong password");
                    continue;
                }
                
            }

            if (cracked)
            {
                MessageBox.Show(finalPassword);
                fileName = fileName + "_cracked.doc";
                FName = fileName;
                aDoc.SaveAs(ref FName, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj);
            }
            aDoc.Close(ref nullobj, ref nullobj, ref nullobj);
            wordlist.Close();
        }

        private void btnOpenWordlist_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                txtWordlist.Text = openFileDialog2.FileName;
                m_wordlistFile = txtWordlist.Text;
                btnCrackPasswd.Enabled = true;
            }
        }
    }

}
