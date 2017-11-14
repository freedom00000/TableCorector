using System;
using System.Windows;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using System.IO;
using System.Collections.Generic;

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

                foreach (var table in docWord.MainDocumentPart.Document.Body.Elements<Table>())
                {
                    List<List<string>> numCells = new List<List<string>>();
                    int i = 0;
                    int prev_shapka = 0;
                    double check;
                    foreach (var item in table.Elements<TableRow>())
                    {

                        if (prev_shapka == 0 ||
                            prev_shapka != item.Elements<TableCell>().Count()
                            && i < 2)
                        {
                            shapkaCorrector(item);
                            i++;
                        }
                        else
                        {
                            int j = 0;
                            List<string> list = new List<string>();

                            foreach (TableCell c in item.Elements<TableCell>())
                            {

                                if (Double.TryParse(c.Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First().Text, out check))
                                {
                                    list.Add(c.Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First().Text);
                                    ParagraphProperties pPr = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                                    c.PrependChild(pPr);
                                }
                                else
                                {
                                    list.Add(null);
                                    ParagraphProperties pPr = new ParagraphProperties(new Justification() { Val = JustificationValues.Both });
                                    c.PrependChild(pPr);
                                }

                                foreach (Paragraph para in c.Elements<Paragraph>())
                                {
                                    foreach (Run run in para.Elements<Run>())
                                    {
                                        RunProperties rPr = new RunProperties(new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                                                      new FontSize() { Val = "24" },
                                                                      new Color() { Val = "000000"});
                                        run.PrependChild(rPr);
                                        run.RemoveChild(run.Elements<RunProperties>().Last());

                                    }
                                }
                                j++;
                            }
                            numCells.Add(list);
                        }

                        prev_shapka = item.Elements<TableCell>().Count();
                    }
                    // numCorrector(numCells, table);
                    i++;
                }
            }

        }
        void shapkaCorrector(TableRow shapka)
        {
            TableRowProperties tbRp = new TableRowProperties(new TableHeader());
            shapka.PrependChild(tbRp);
            foreach (TableCell cell in shapka.Elements<TableCell>())
            {
                ParagraphProperties pPr = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                cell.PrependChild(pPr);
                cell.RemoveChild(cell.Elements<ParagraphProperties>().Last());

                foreach (Paragraph para in cell.Elements<Paragraph>())
                {
                    foreach (Run run in para.Elements<Run>())
                    {
                        RunProperties rPr = new RunProperties(new RunFonts() { Ascii = "Times New Roman" , HighAnsi = "Times New Roman" },
                                                      new FontSize() { Val = "28" },
                                                      new Color() { Val = "000000" });
                        run.PrependChild(rPr);
                        run.RemoveChild(run.Elements<RunProperties>().Last());
                    }
                }
            }
        }


        void numCorrector(List<List<string>> list, Table table)
        {
            int count = 0;
            for (int i = 0; i < list[1].Count; i++)
            {
                for (int j = 0; j < list.Count; j++)
                {
                    if (list[j][i] != null &&
                        count < list[j][i].Split(',')[0].Count())
                    {
                        count = list[j][i].Split(',')[0].Count();
                    }
                }

                for (int j = 0; j < list.Count; j++)
                {
                    if (list[j][i] != null)
                    {
                        list[j][i] = addSpace(count, list[j][i]);
                    }
                }
            }

            insertTable(table, list);
        }

        string addSpace(int count, string str)
        {
            int str_count = 0;

            str_count = count - str.Split(',')[0].Count();
            str = str.Insert(0, new string(Char.Parse("\u00A0"), str_count * 2));


            return str;
        }

        void insertTable(Table table, List<List<string>> list)
        {
            int j = -1;
            int i = -1;
            foreach (var item in table.Elements<TableRow>())
            {
                j = 0;
                foreach (TableCell c in item.Elements<TableCell>())
                {

                    foreach (Paragraph para in c.Elements<Paragraph>())
                    {
                        foreach (Run run in para.Elements<Run>())
                        {

                            foreach (var t in run.Elements<Text>())
                            {
                                if (i >= 0 && j >= 0 && j < 4)
                                    if (list[i][j] != null)
                                        t.Text = list[i][j];
                            }

                        }

                    }
                    j++;
                }

                i++;

            }

        }

    }
}
