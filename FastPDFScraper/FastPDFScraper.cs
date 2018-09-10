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

namespace FastPDFScraper
{
    public partial class FastPDFScraper : Form
    {
        private string[] pdfFiles = null;
        private List<string> protectedFiles = new List<string>();
        private string keyFile = null;
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
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK)
            {
                pdfFiles = Directory.GetFiles(fbd.SelectedPath);
            }
            if (pdfFiles == null) return;
            totalPages = 0;
            for(int i = 0;i < pdfFiles.Length; i++)
            {
                try
                {
                    PdfReader reader = new PdfReader(pdfFiles[i]);
                    totalPages += reader.NumberOfPages;
                }catch(Exception ex)
                {
                    protectedFiles.Add(pdfFiles[i]);
                }
            }
            progressBar.Maximum = totalPages;
            currentPage = 0;
        }

        private void btnOpenKeywordFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "CSV Files|*.csv";
            DialogResult result = ofd.ShowDialog();

            if (result == DialogResult.OK)
            {
                keyFile = ofd.FileName;
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {           
                currentPage = 0;
                BarValue = 0;
                if (pdfFiles == null || pdfFiles.Length == 0 || keyFile == null) return;
                string text = System.IO.File.ReadAllText(keyFile);
                string[] keys = text.Split(',');
                for (int i = 0; i < pdfFiles.Length; i++)
                {
                    PdfReader reader = new PdfReader(pdfFiles[i]);
                    Console.WriteLine("i=" + i);
                    for (int j = 1; j <= reader.NumberOfPages; j++)
                    {
                        Console.WriteLine("j=" + j);
                        string pdf = PdfTextExtractor.GetTextFromPage(reader, j);
                        Parameter param = new Parameter()
                        {
                            Keys = keys,
                            FileName = pdfFiles[i],
                            Page = j,
                            PDF = pdf
                        };
                        Thread thread = new Thread(new ParameterizedThreadStart(search));
                        thread.Start(param);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("PDF Open Exception");
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
                MessageBox.Show("PDF Search Exception");
            }
        }

        void printResult()
        {
            List<Record> newResult = result.OrderBy(c => c.Key).ThenBy(c => c.FileName).ThenBy(c => c.Page).ToList();
            using (StreamWriter writetext = new StreamWriter("result.csv"))
            {
                writetext.WriteLine("Key,FileName,PageNo");

                var grouped = newResult.GroupBy(p => new { p.Key, p.FileName });
                foreach (var group in grouped)
                {
                    string record = group.Key.Key + "," + Path.GetFileName(group.Key.FileName) + ",";
                    if (chkMultiResult.Checked)
                    {
                        foreach (var product in group)
                        {
                            record = record + product.Page + " | ";
                        }
                    }
                    else
                    {
                        record = record + group.First().Page;
                    }
                    writetext.WriteLine(record);
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
}
