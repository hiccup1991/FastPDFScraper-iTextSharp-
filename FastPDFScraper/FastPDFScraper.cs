using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Text.RegularExpressions;

namespace FastPDFScraper
{
    public partial class FastPDFScraper : Form
    {
        private string[] pdfFiles = null;
        private List<string> readableFiles = new List<string>();
        private List<string> protectedFiles = new List<string>();
        private string keyFile = null;
        private string[] keys = null;
        List<Record> result = new List<Record>();
        private int totalPages = 0;
        private int currentPage = 0;

        public int BarValue
        {
            set
            {
                if (progressBar.InvokeRequired)
                {
                    progressBar.Invoke((MethodInvoker)delegate
                    {
                        BarValue = value;
                    });
                }
                else
                {
                    progressBar.Value = value;
                    progressBar.Invalidate();
                    progressBar.Update();
                }
            }
        }

        public FastPDFScraper()
        {
            InitializeComponent();
        }

        private void btnOpenPDFFolder_Click(object sender, EventArgs e)
        {
            
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult dlg = fbd.ShowDialog();
            if (dlg == DialogResult.OK)
            {
                pdfFiles = Directory.GetFiles(fbd.SelectedPath);
            }
            if (pdfFiles == null) return;
            totalPages = 0;
            readableFiles.Clear();
            protectedFiles.Clear();
            progressBar.Maximum = pdfFiles.Length;
            progressReport.Visible = true;
            progressReport.Text = "Opening Files...";
            progressReport.Update();
            for(int i = 0;i < pdfFiles.Length; i++)
            {
                BarValue = i + 1;
                try
                {
                    PdfReader reader = new PdfReader(pdfFiles[i]);
                    totalPages += reader.NumberOfPages;
                    readableFiles.Add(pdfFiles[i]);
                }catch(Exception ex)
                {
                    protectedFiles.Add(pdfFiles[i]);
                }
            }
            if(protectedFiles.Count > 0)
                progressReport.Text = protectedFiles.Count + " Files Open Failed.";
            else
                progressReport.Text = "All Files Open";
            currentPage = 0;
            progressReport.Update();
        }

        private void btnOpenKeywordFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "CSV Files|*.csv";
            DialogResult dlg = ofd.ShowDialog();

            if (dlg == DialogResult.OK)
            {
                keyFile = ofd.FileName;
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            BarValue = 0;
            progressBar.Maximum = totalPages;
            currentPage = 0;
            result.Clear();
            progressReport.Visible = false;
            if (readableFiles == null || readableFiles.Count == 0 || keyFile == null) return;
            string text = System.IO.File.ReadAllText(keyFile);
            text.Trim();
            keys = Regex.Split(text, "\r\n|\r|\n");
            List<string> temp = keys.ToList();
            temp.Remove("Keys");
            temp.Remove("");
            keys = temp.ToArray();
            for (int i = 0; i < readableFiles.Count; i++)
            {
                try
                {
                    PdfReader reader = new PdfReader(readableFiles[i]);
                    Console.WriteLine("i=" + i);
                    for (int j = 1; j <= reader.NumberOfPages; j++)
                    {
                        Console.WriteLine("j=" + j);
                        string pdf = PdfTextExtractor.GetTextFromPage(reader, j);
                        Parameter param = new Parameter()
                        {
                            Keys = keys,
                            FileName = readableFiles[i],
                            Page = j,
                            PDF = pdf
                        };
                        Thread thread = new Thread(new ParameterizedThreadStart(search));
                        thread.Start(param);
                    }
                }
                catch(Exception ex)
                {
                    //protectedFiles.Add(readableFiles[i]);
                    MessageBox.Show("Searching error on " + readableFiles[i]);
                }
            }
        }

        void search(object param)
        {
            try
            {            
                Parameter p = (Parameter)param;
                string Text = p.PDF;
                //MessageBox.Show(Text);
                if (Text != null)
                {
                    for (int i = 0; i < p.Keys.Length; i++)
                        if (Text.Contains(p.Keys[i]))
                            result.Add(new Record() { Key = p.Keys[i], FileName = p.FileName, Page = p.Page });
                }
                currentPage++;
                BarValue = currentPage;
            } catch (Exception ex)
            {
                //MessageBox.Show("PDF Search Exception");
            }
        }

        void printResult()
        {
            /*foreach (var key in keys)
            {
                foreach (var filename in readableFiles)
                {
                    if( !result.Exists(x => (x.Key == key) && (x.FileName == filename)))
                        result.Add(new Record() { Key = key, FileName = filename, Page = -1});
                }
            }*/
            List<Record> newResult = result.OrderBy(c => c.Key).ThenBy(c => c.FileName).ThenBy(c => c.Page).ToList();
            
            using (StreamWriter writetext = new StreamWriter("result.csv"))
            {
                writetext.WriteLine("Key,FileName,PageNo,Comment");

                /*var grouped = newResult.GroupBy(p => new { p.Key, p.FileName });
                foreach (var group in grouped)
                {
                    string record = group.Key.Key + "," + Path.GetFileName(group.Key.FileName) + ",";
                    if (chkMultiResult.Checked)
                    {
                        foreach (var product in group)
                        {
                            if (product.Page != -1)
                                record = record + product.Page + " | ";
                            else
                                record = record + " ,key not found";
                        }
                        if (!record.Contains("key not found")) record = record + ",key found";
                    }
                    else
                    {
                        if (group.First().Page != -1)
                            record = record + group.First().Page + ",key found";
                        else
                            record = record + " ,key not found";
                    }

                    writetext.WriteLine(record);
                }*/

                List<OutputRecord> outputTempResult = new List<OutputRecord>();
                var grouped = newResult.GroupBy(p => new { p.Key, p.FileName });
                foreach (var group in grouped)
                {
                    OutputRecord record = new OutputRecord() { Key = group.Key.Key, FileNames = Path.GetFileName(group.Key.FileName) };
                    string pages = "";
                    if (chkMultiResult.Checked)
                    {
                        foreach (var product in group)
                        {
                            pages = pages + product.Page + " . ";
                        }
                        pages = pages.Remove(pages.Length - 3);
                    }
                    else
                    {
                        pages = group.First().Page + "";
                    }
                    record.Pages = pages;
                    outputTempResult.Add(record);
                }

                List<OutputRecord> outputResult = new List<OutputRecord>();
                var newgrouped = outputTempResult.GroupBy(p => new { p.Key });
                foreach (var group in newgrouped)
                {
                    OutputRecord record = new OutputRecord() { Key = group.Key.Key };
                    string pages = "";
                    string filenames = "";
                    foreach (var product in group)
                    {
                        pages = pages + product.Pages + " | ";
                        filenames = filenames + product.FileNames + " | ";
                    }
                    record.Pages = pages.Remove(pages.Length - 3);
                    record.FileNames = filenames.Remove(filenames.Length - 3);
                    outputResult.Add(record);
                }

                foreach (var key in keys)
                {
                    if(!outputResult.Exists(x => x.Key == key)){
                        writetext.WriteLine(key + ", , ,key not found");
                    }
                         //outputResult.Add(new Record() { Key = key, FileName = filename, Page = -1});
                    else
                    {
                        OutputRecord temp = outputResult.Find(x => x.Key == key);
                        writetext.WriteLine(temp.Key + "," + temp.FileNames+ "," + temp.Pages + ",key found");
                    }

                }
            }

            using (StreamWriter writefilelist = new StreamWriter("protectedFilesList.csv"))
            {
                for (int i = 0; i < protectedFiles.Count; i++)
                {
                    writefilelist.WriteLine(protectedFiles[i]);
                }
            }
            MessageBox.Show("Printing Complete!");
        }

        private void btnPrintResult_Click(object sender, EventArgs e)
        {
            printResult();
        }
    }
    public class Parameter
    {
        public string[] Keys { get; set; }
        public string FileName { get; set; }
        public int Page { get; set; }
        public string PDF { get; set; }
    }

    public class Record : IEquatable<Record>
    {
        public string Key { get; set; }
        public string FileName { get; set; }
        public int Page { get; set; }

        public bool Equals(Record other)
        {
            if (other == null) return false;
            else if (other.Key == this.Key && other.FileName == this.FileName && other.Page == this.Page) return true;
            return false; 
        }
    }

    public class OutputRecord
    {
        public string Key { get; set; }
        public string FileNames { get; set; }
        public string Pages { get; set; }
    }
}
