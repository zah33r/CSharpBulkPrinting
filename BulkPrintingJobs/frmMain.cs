using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace BulkPrinter
{
    public partial class frmMain : Form
    {
        #region Local Variables

        private string imagePath = string.Empty;

        #endregion

        #region ctor

        public frmMain()
        {
            InitializeComponent();
        }

        #endregion

        #region Private Methods

        private void SendWordToPrinter(string path)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application wordInstance = new Microsoft.Office.Interop.Word.Application();
                FileInfo wordFile = new FileInfo(path);
                object fileObject = wordFile.FullName;
                object oMissing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Document doc = wordInstance.Documents.Open(ref fileObject, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                doc.Activate();
                doc.PrintOut(oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                wordInstance.Quit();
            }
            catch (Exception exp)
            {
                string errorMsg = exp.Message;
            }
        }

        private void SendToPrinter(string path)
        {
            try
            {
                ProcessStartInfo info = new ProcessStartInfo();
                info.Verb = "print";
                info.FileName = path;
                info.CreateNoWindow = true;
                info.WindowStyle = ProcessWindowStyle.Hidden;

                Process p = new Process();
                p.StartInfo = info;
                p.Start();

                if (path.ToLower().Trim().EndsWith(".pdf"))
                {
                    p.WaitForInputIdle(3000);
                    System.Threading.Thread.Sleep(3000);

                    if (p != null)
                        p.Kill();
                }
                else
                    p.WaitForExit();

                System.Threading.Thread.Sleep(5000);
            }
            catch (Exception exp)
            {
                string errorMsg = exp.Message;
            }
        }

        private void listBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                string pname = this.listBox1.SelectedItem.ToString();
                myPrinters.SetDefaultPrinter(pname);
            }
            catch (Exception exp)
            {
                string errorMsg = exp.Message;
            }
        }

        private void listAllPrinters()
        {
            foreach (var item in PrinterSettings.InstalledPrinters)
            {
                this.listBox1.Items.Add(item.ToString());
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                listAllPrinters();
            }
            catch (Exception exp)
            {
                string errorMsg = exp.Message;
                MessageBox.Show("An error occured while retreiving the list of printers" + Environment.NewLine + Environment.NewLine + "Error Details: " + exp.Message);
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (DialogResult.OK == folderBrowserDialog1.ShowDialog())
                {
                    #region Filters

                    dataGridView1.DataSource = null;
                    List<string> listOfFiles = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.doc", SearchOption.TopDirectoryOnly).ToList<string>();  //docx as well
                    listOfFiles.AddRange(Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.pdf", SearchOption.TopDirectoryOnly).ToList<string>());
                    listOfFiles.AddRange(Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.gif", SearchOption.TopDirectoryOnly).ToList<string>());
                    listOfFiles.AddRange(Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.jpg", SearchOption.TopDirectoryOnly).ToList<string>());
                    listOfFiles.AddRange(Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.xlsx", SearchOption.TopDirectoryOnly).ToList<string>());
                    
                    #endregion

                    UpdateGrid(listOfFiles);
                }
            }
            catch (Exception exp)
            {
                string errorMsg = exp.Message;
            }
        }

        private void UpdateGrid(List<string> listOfFiles)
        {
            try
            {
                DataTable table = new DataTable("Files");
                table.Columns.Add("FileName", typeof(string));
                table.Columns[0].Caption = "File Name";
                listOfFiles.ForEach((file) => { table.Rows.Add(file); });
                dataGridView1.DataSource = table;
                dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView1.Columns[0].HeaderText = "File Name";
                dataGridView1.Update();
            }
            catch (Exception Exp)
            {
                string errorMsg = Exp.Message;
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.DataSource == null)
                {
                    MessageBox.Show("Please select a file or directory to proceed.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }


                if (DialogResult.Yes == MessageBox.Show("Are you sure, you want to print all files?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                {

                    foreach (DataGridViewRow item in dataGridView1.Rows)
                    {
                        string path = item.Cells[0].Value.ToString();
                        item.Cells[0].Selected = true;
                        try
                        {
                            if (path.ToLower().Trim().EndsWith(".jpg") || path.ToLower().Trim().EndsWith(".jpeg") || path.ToLower().Trim().EndsWith(".gif"))
                            {
                                this.imagePath = path;
                                printDocument1.Print();
                            }
                            else if (path.ToLower().Trim().EndsWith(".doc") || path.ToLower().Trim().EndsWith(".docx"))
                            {
                                SendWordToPrinter(path);
                            }
                            else
                                SendToPrinter(path);
                        }
                        catch (Exception innerExp)
                        {
                            string errorMsg = innerExp.Message;
                        }
                        //Sleep the process for 5 seconds in case files are still getting skipped
                        //System.Threading.Thread.Sleep(5000);
                        item.Cells[0].Selected = false;
                    }

                    MessageBox.Show("All Jobs were sent to printer successfully.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception exp)
            {
                string errorMsg = exp.Message;
            }
        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            try
            {
                System.Drawing.Image imageObject = System.Drawing.Image.FromFile(imagePath);
                Point location = new Point(100, 100);
                e.Graphics.DrawImage(imageObject, location);
            }
            catch (Exception exp)
            {
                string errorMsg = exp.Message;
            }
        }        

        private void selectFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (DialogResult.OK == openFileDialog1.ShowDialog())
                {
                    UpdateGrid(openFileDialog1.FileNames.ToList<string>());
                }
            }
            catch (Exception exp)
            {
                string errorMsg = exp.Message;
            }
        }

        #endregion

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }

    public static class myPrinters
    {
        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string Name);

    }
}
