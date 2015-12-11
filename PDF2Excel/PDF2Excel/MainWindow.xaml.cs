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
        public class PDFTextExtractor
        {
            /// BT = Beginning of a text object operator 
            /// ET = End of a text object operator
            /// Td move to the start of next line
            ///  5 Ts = superscript
            /// -5 Ts = subscript

            #region Fields

            #region _numberOfCharsToKeep
            /// <summary>
            /// The number of characters to keep, when extracting text.
            /// </summary>
            private static int _numberOfCharsToKeep = 15;
            #endregion

            #endregion



            #region ExtractTextFromPDFBytes
            /// <summary>
            /// This method processes an uncompressed Adobe (text) object 
            /// and extracts text.
            /// </summary>
            /// <param name="input">uncompressed</param>
            /// <returns></returns>
            public string ExtractTextFromPDFBytes(byte[] input)
            {
                if (input == null || input.Length == 0) return "";

                try
                {
                    string resultString = "";

                    // Flag showing if we are we currently inside a text object
                    bool inTextObject = false;
                    // Flag showing if the next character is literal 
                    // e.g. '\\' to get a '\' character or '\(' to get '('
                    bool nextLiteral = false;

                    // () Bracket nesting level. Text appears inside ()
                    int bracketDepth = 0;

                    // Keep previous chars to get extract numbers etc.:
                    char[] previousCharacters = new char[_numberOfCharsToKeep];
                    for (int j = 0; j < _numberOfCharsToKeep; j++) previousCharacters[j] = ' ';


                    for (int i = 0; i < input.Length; i++)
                    {
                        char c = (char)input[i];
                        if (inTextObject)
                        {
                            // Position the text
                            if (bracketDepth == 0)
                            {

                                if (CheckToken(new string[] { "TD", "Td" }, previousCharacters))
                                {
                                    resultString += "\n\r";
                                }
                                else
                                {
                                    if (CheckToken(new string[] { "'", "T*", "\"" }, previousCharacters))
                                    {
                                        resultString += "\n";
                                    }
                                    else
                                    {
                                        if (CheckToken(new string[] { "Tj" }, previousCharacters))
                                        {
                                            resultString += " ";
                                        }
                                    }
                                }
                            }

                            // End of a text object, also go to a new line.
                            if (bracketDepth == 0 &&
                                CheckToken(new string[] { "ET" }, previousCharacters))
                            {

                                inTextObject = false;
                                resultString += " ";
                            }
                            else
                            {
                                // Start outputting text
                                if ((c == '(') && (bracketDepth == 0) && (!nextLiteral))
                                {
                                    bracketDepth = 1;
                                }
                                else
                                {
                                    // Stop outputting text
                                    if ((c == ')') && (bracketDepth == 1) && (!nextLiteral))
                                    {
                                        bracketDepth = 0;
                                    }
                                    else
                                    {
                                        // Just a normal text character:
                                        if (bracketDepth == 1)
                                        {

                                            // Only print out next character no matter what. 
                                            // Do not interpret.
                                            if (c == '\\' && !nextLiteral)
                                            {

                                                nextLiteral = false;
                                            }
                                            else
                                            {
                                                
                                                if (((c >= ' ') && (c <= '~')) ||
                                                    ((c >= 128) && (c < 255)))
                                                {
                                                    if (resultString.Contains("040"))
                                                    {
                                                        resultString.Replace("040", " ");
                                                    }
                                                    resultString += c.ToString();
                                                }
                                                nextLiteral = false;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // Store the recent characters for 
                        // when we have to go back for a checking
                        for (int j = 0; j < _numberOfCharsToKeep - 1; j++)
                        {
                            previousCharacters[j] = previousCharacters[j + 1];
                        }
                        previousCharacters[_numberOfCharsToKeep - 1] = c;

                        // Start of a text object
                        if (!inTextObject && CheckToken(new string[] { "BT" }, previousCharacters))
                        {
                            inTextObject = true;
                        }
                    }
                    return resultString;
                    
                }
                catch
                {
                    return "";
                }
            }
            #endregion

            #region CheckToken
            /// <summary>
            /// Check if a certain 2 character token just came along (e.g. BT)
            /// </summary>
            /// <param name="search">the searched token</param>
            /// <param name="recent">the recent character array</param>
            /// <returns></returns>
            private bool CheckToken(string[] tokens, char[] recent)
            {
                foreach (string token in tokens)
                {
                    if (token.Length > 1)
                    {
                        if ((recent[_numberOfCharsToKeep - 3] == token[0]) &&
                            (recent[_numberOfCharsToKeep - 2] == token[1]) &&
                            ((recent[_numberOfCharsToKeep - 1] == ' ') ||
                            (recent[_numberOfCharsToKeep - 1] == 0x0d) ||
                            (recent[_numberOfCharsToKeep - 1] == 0x0a)) &&
                            ((recent[_numberOfCharsToKeep - 4] == ' ') ||
                            (recent[_numberOfCharsToKeep - 4] == 0x0d) ||
                            (recent[_numberOfCharsToKeep - 4] == 0x0a))
                            )
                        {
                            return true;
                        }
                    }
                    else
                    {
                        return false;
                    }

                }
                return false;
            }
            #endregion
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
                        outputText = new PDFTextExtractor().ExtractTextFromPDFBytes(stream.Value);

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
                        Console.WriteLine(fullText[i].Replace("040", " ".Replace("050", "(".Replace("051", ")"))));
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
