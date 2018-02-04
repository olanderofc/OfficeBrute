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

        bool m_cancel = false;
        UInt64 elapsed_time = 0;

        UInt64 m_wordlistRows = 0;
        UInt64 THREADS = 4;
        UInt64 m_first = 0;
        UInt64 m_second = 0;
        UInt64 m_third = 0;
        UInt64 m_fourth = 0;
        
        public Form1()
        {
            InitializeComponent();
        }
        private void SetText(string text)
        {
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
                txtDocfile.Text = openFileDialog1.FileName;
            }
        }
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            if (txtDocfile.Text != null && txtWordlist != null)
            {

                m_cancel = false;
                m_fileName = txtDocfile.Text;
                btnStopCracking.Enabled = true;
                btnCrackPasswd.Enabled = false;
                start();
                label1.Visible = true;
                label2.Visible = true;
                timer1.Interval = 1000;
                timer1.Enabled = true;
            }
        }



         /*
            You can modify this function to scale up and use more threads. Right now it is only 4 threads.
         */
        public void start()
        {

            m_wordlistRows = (UInt64)File.ReadLines(m_wordlistFile).Count();
            //Rows is 23403 / 4 =  5850
            UInt64 index = m_wordlistRows / THREADS;
            UInt64 rest = m_wordlistRows % THREADS;

            m_first = 0;
            m_second = m_first + index; // 5850
            m_third = m_second + index; // 5850 + 5850
            m_fourth = m_third + index + rest; // 5850 + 5850 + 5850
            while (m_fourth > m_wordlistRows)
            {
                m_fourth--;
            }

          
            try
            {
                File.Copy(m_fileName, m_fileName + "_1.doc");
                File.Copy(m_fileName, m_fileName + "_2.doc");
                File.Copy(m_fileName, m_fileName + "_3.doc");
                File.Copy(m_fileName, m_fileName + "_4.doc");
            }
            catch (Exception)
            {
                MessageBox.Show("Error creating temporary files. Make sure you are able to write files in the same directory as the application", "File copy failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Thread t = new Thread(new ParameterizedThreadStart(threadBruteForce));
            UInt64 end = m_first + index - 1;
            BruteObj a = new BruteObj(m_fileName + "_1.doc", m_wordlistFile, m_first, end);
            t.Start((object)a);

            Thread u = new Thread(new ParameterizedThreadStart(threadBruteForce));
            end = m_second + index - 1;
            BruteObj b = new BruteObj(m_fileName + "_2.doc", m_wordlistFile, m_second, end);
            u.Start((object)b);

            Thread r = new Thread(new ParameterizedThreadStart(threadBruteForce));
            end = m_third + index - 1;
            BruteObj c = new BruteObj(m_fileName + "_3.doc", m_wordlistFile, m_third, end);
            r.Start((object)c);

            Thread s = new Thread(new ParameterizedThreadStart(threadBruteForce));
            end = m_fourth + index - 1;
            BruteObj d = new BruteObj(m_fileName + "_4.doc", m_wordlistFile, m_fourth, end);       
            s.Start((object)d);

        }
       

       public void threadBruteForce(object bruteObj)
       {

           BruteObj brute = (BruteObj)bruteObj;
           string fileName = brute.getPath();
           string wordlistFile = brute.getWordlistPath();
           System.IO.StreamReader wordlist = new System.IO.StreamReader(wordlistFile);

           // Setup Microsoft Word connection
           Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
           object nullobj = System.Reflection.Missing.Value;
           object password = "tmp";
           object FName = fileName;
           // https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.document.aspx
           Microsoft.Office.Interop.Word.Document aDoc = null;
           wordApp.Visible = false;
           // Open the document
           aDoc = wordApp.Documents.Open(ref FName, ref nullobj, ref nullobj, ref nullobj, ref password, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj);
           // WARNING - This will run document if it is using macros. You can modify this if you want.
           // aDoc.Activate();

           //Check if the document is protected
           if (aDoc.ProtectionType == Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
           {
               MessageBox.Show("Document is not protected. Will not proceed");
               timer1.Enabled = false;
               lblTime.Text = "";
               return;   
           }

           string finalPassword = "";
           bool cracked = false;
           string line;
           UInt64 counter = 0;
           UInt64 start = brute.getStart();
           UInt64 end = brute.getEnd();
           // Do not read in the entire file in memory as it will be very slow.
           while ((line = wordlist.ReadLine()) != null)
           {

               if (counter > end)
               {
                   break;
               }

               if (m_cancel)
               {
                   break;
               }

               if (counter >= start && counter <= end)
               {
                   try
                   {
                       this.SetText(line);
                       object passwd = line;
                       aDoc.Unprotect(ref passwd);
                       timer1.Enabled = false;
                       MessageBox.Show("Success. Password is: '" + line + "'.Cracking took " + lblTime.Text + " Seconds", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                       btnStopCracking.Text = "Restart";
                       finalPassword = line;
                       cracked = true;
                       break;
                   }
                   catch (Exception)
                   {
                       counter++;
                       continue;
                   }
               }
               else
               {
                   counter++;
               }

              
           }

           if (cracked)
           {
               try
               {
                   string newFileName = fileName + "_cracked.doc";
                   object newFName = newFileName;
                   aDoc.SaveAs(ref newFName, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj, ref nullobj);

               }
               catch (Exception)
               {
                   MessageBox.Show("Unable to save cracked document");
               }
           }
           //Close the document
           aDoc.Close(ref nullobj, ref nullobj, ref nullobj);
           wordlist.Close();

           //Cleanup
           try
           {
               File.Delete(fileName);
           }
           catch (Exception)
           {
               MessageBox.Show("Unable to delete temporary file " + fileName + ". Make sure file is not in use", "Delete temporary file", MessageBoxButtons.OK, MessageBoxIcon.Error);
           }
           
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

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            label1.Visible = false;
            label2.Visible = false;
            timer1.Enabled = false;
            lblCurrentPass.Text = "";
            lblTime.Text = "";
            m_cancel = true;
            MessageBox.Show("Cracking stopped. Make sure you kill any running Word Processes",
                "Cracking Stopped", MessageBoxButtons.OK,MessageBoxIcon.Warning);
            btnCrackPasswd.Enabled = true;
            btnStopCracking.Enabled = false;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            elapsed_time += 1;
            lblTime.Text = Convert.ToString(elapsed_time);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("OfficeBrute v0.2 by @olanderofc. Brute force office protection passwords. Use on your own risk", "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

}
