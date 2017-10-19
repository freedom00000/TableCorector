using System;
using System.Windows;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using System.IO;

namespace TableCorector
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        string filepath = "test.docx";

        private void button_Click(object sender, RoutedEventArgs e)
        {
            DirectoryInfo di = new DirectoryInfo(".");
            foreach (var file in di.GetFiles("*.docx"))
                file.Delete();
            DirectoryInfo di2 = new DirectoryInfo("../../../");
            foreach (var file in di2.GetFiles("*.docx"))
                file.CopyTo(di.FullName + "/" + file.Name);

            using (WordprocessingDocument docWord = WordprocessingDocument.Open(filepath, true))
            {
                Table table1 = docWord.MainDocumentPart.Document.Body.Elements<Table>().First();
                foreach (var table in docWord.MainDocumentPart.Document.Body.Elements<Table>())
                {

                    int i = 0;
                    int check;
                    foreach (var item in table.Elements<TableRow>())
                    {
                        if (i == 0)
                        {
                            shapkaCorrector(item);
                            i++;
                        }
                        else
                        {
                            foreach (TableCell c in item.Elements<TableCell>())
                            {
                                if (Int32.TryParse(c.Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First().Text, out check))
                                {
                                    ParagraphProperties pPr = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                                    c.PrependChild(pPr);
                                }
                                else
                                {
                                    ParagraphProperties pPr = new ParagraphProperties(new Justification() { Val = JustificationValues.Both });
                                    c.PrependChild(pPr);
                                }

                                foreach (Paragraph para in c.Elements<Paragraph>())
                                {
                                    foreach (Run run in para.Elements<Run>())
                                    {
                                        RunProperties rPr = new RunProperties(new RunFonts() { Ascii = "Times New Roman" },
                                                                      new FontSize() { Val = "24" });
                                        run.PrependChild(rPr);
                                    }
                                }
                            }

                        }

                    }
                }
            }

        }
        void shapkaCorrector(TableRow shapka)
        {
            foreach (TableCell cell in shapka.Elements<TableCell>())
            {
                ParagraphProperties pPr = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                cell.PrependChild(pPr);

                foreach (Paragraph para in cell.Elements<Paragraph>())
                {
                    foreach (Run run in para.Elements<Run>())
                    {
                        RunProperties rPr = new RunProperties(new RunFonts() { Ascii = "Times New Roman" },
                                                      new FontSize() { Val = "28" });
                        run.PrependChild(rPr);
                    }
                }
            }
        }
    }
}
