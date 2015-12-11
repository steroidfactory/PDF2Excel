using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Globalization;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using System.IO;

namespace PDF2Excel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void test(object sender, RoutedEventArgs e)
        {
            //@"e:\WIP\5836015.pdf"

            PDFParser();
        }
        

        public void PDFParser()
        {
            var streamWriter = new StreamWriter(@"e:\WIP\5836015.txt", false);

            String outputText = "";

         //   try
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
                    }

                }

            }
          //  catch (Exception e)
          //  {

         //   }
          //  streamWriter.Close();
        }
    }
}
