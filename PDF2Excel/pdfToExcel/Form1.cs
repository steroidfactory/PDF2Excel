using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using System.IO;
using PDF2Excel;

namespace pdfToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public void PDFParser()
        {
            var streamWriter = new StreamWriter(@"e:\WIP\5836015.txt", false);

            String outputText = "";

            try
            {
                PdfDocument inputDocument = PdfReader.Open(@"e:\WIP\5836015.pdf", PdfDocumentOpenMode.ReadOnly);
                List<string> fullText = new List<string>();
                int lineCount = 0;
                string tempLine = "";
                List<char> exceptionList = new List<char>();
                //exceptionList.Add((char)32);
                exceptionList.Add((char)13);
                exceptionList.Add((char)10);
                foreach (PdfPage page in inputDocument.Pages)
                {
                    for (int index = 0; index < page.Contents.Elements.Count; index++)
                    {
                        PdfDictionary.PdfStream stream = page.Contents.Elements.GetDictionary(index).Stream;
                        outputText = new PDF().ExtractTextFromPDFBytes(stream.Value);

                    }
                    foreach (char ch in outputText)
                    {
                        if (!exceptionList.Contains(ch))
                        {
                            tempLine += ch.ToString();
                        }
                        if (ch == '\n')
                        {
                            if (tempLine != "")
                            {
                                fullText.Add(tempLine);
                            }
                            tempLine = "";
                            lineCount++;
                        }
                    }

                    for (int i = 0; i < fullText.Count - 1; i++)
                    {
                        Console.WriteLine(fullText[i].Replace("040", " ").Replace("050", "(").Replace("051", ")"));
                        streamWriter.WriteLine(fullText[i].Replace("040", " ").Replace("050", "(").Replace("051", ")"));
                    }

                }

            }
            catch (Exception e)
            {

            }
            streamWriter.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PDFParser();
        }
    }
}
