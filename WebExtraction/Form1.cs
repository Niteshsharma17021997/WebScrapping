using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http;
using System.Configuration;

namespace WebExtraction
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string url = textBox1.Text;
            HtmlAgilityPack.HtmlWeb web = new HtmlAgilityPack.HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = web.Load(url);
            List<string> field = new List<string>();
            List<string> value = new List<string>();
            
            var headers = doc.DocumentNode.Descendants("h3");
            var values = doc.DocumentNode.Descendants("p");
            foreach (var h in headers)
            {
                field.Add(h.InnerHtml.ToString().Replace("&bull;","*").Replace("<br>","  "));
            }
            Console.WriteLine();
            foreach (var p in values)
            {
                value.Add(p.InnerHtml.ToString().Replace("&bull;", "*").Replace("<br>", "  "));
            }
            string rootFolder = ConfigurationManager.AppSettings["path"];
            string fileName = @"Book.xlsx";
            FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
            bool insert = false;
            if (file.Exists)
            {
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet workSheet = package.Workbook.Worksheets["Sheet1"];
                    int rowCount = 1;
                    if (workSheet.Dimension != null)
                    {
                        rowCount = workSheet.Dimension.End.Row;
                    }
                    for (int i = 1; i <= rowCount;)
                    {
                        if (workSheet.Cells[i, 1].Value != null)
                        {
                            if (workSheet.Cells.Last(c => c.Start.Row == i).End.Column == field.Count)
                            {
                                int j = 1;
                                for (j = 1; j <= field.Count; j++)
                                {
                                    if (workSheet.Cells[i, j].Value.ToString() != field[j - 1])
                                    {
                                        while (workSheet.Cells[i, 1].Value != null)
                                        {
                                            i++;
                                        }
                                        break;
                                    }
                                }
                                if (j == field.Count + 1)
                                {
                                    while (workSheet.Cells[i, 1].Value != null)
                                    {
                                        i++;
                                    }
                                    workSheet.InsertRow(i, 1);
                                    for (int m = 1; m <= field.Count; m++)
                                    {
                                        workSheet.Cells[i, m].Value = value[m - 1];
                                    }
                                    insert = true;
                                    break;
                                }
                            }
                            else { i++; }
                        }
                        else { i++; }
                    }

                    for (int i = 1; i <= field.Count && !insert; i++)
                    {
                        if (rowCount == 1)
                        {
                            workSheet.Cells[1, i].Value = field[i - 1];
                            workSheet.Cells[2, i].Value = value[i - 1];
                        }
                        else
                        {
                            workSheet.Cells[rowCount + 2, i].Value = field[i - 1];
                            workSheet.Cells[rowCount + 3, i].Value = value[i - 1];
                        }
                    }
                    package.Save();
                }
            }
            else
            {
                using (ExcelPackage excelPackage = new ExcelPackage())
                {

                    //Create the WorkSheet
                    ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                    int rowCount = 1;
                    if (workSheet.Dimension != null)
                    {
                        rowCount = workSheet.Dimension.End.Row;
                    }
                    for (int i = 1; i <= rowCount;)
                    {
                        if (workSheet.Cells[i, 1].Value != null)
                        {
                            if (workSheet.Cells.Last(c => c.Start.Row == i).End.Column == field.Count)
                            {
                                int j = 1;
                                for (j = 1; j < field.Count; j++)
                                {
                                    if (workSheet.Cells[i, j].Value.ToString() != field[j - 1])
                                    {
                                        while (workSheet.Cells[i, 1].Value != null)
                                        {
                                            i++;
                                        }
                                        break;
                                    }
                                }
                                if (j == field.Count)
                                {
                                    while (workSheet.Cells[i, 1].Value != null)
                                    {
                                        i++;
                                    }
                                    workSheet.InsertRow(i, 1);
                                    for (int m = 1; m <= field.Count; m++)
                                    {
                                        workSheet.Cells[i, m].Value = value[m - 1];
                                    }
                                    insert = true;
                                    break;
                                }
                            }
                            else { i++; }
                        }
                        else { i++; }
                    }

                    for (int i = 1; i <= field.Count && !insert; i++)
                    {
                        if (rowCount == 1)
                        {
                            workSheet.Cells[1, i].Value = field[i - 1];
                            workSheet.Cells[2, i].Value = value[i - 1];
                        }
                        else
                        {
                            workSheet.Cells[rowCount + 2, i].Value = field[i - 1];
                            workSheet.Cells[rowCount + 3, i].Value = value[i - 1];
                        }
                    }
                    FileInfo fi = new FileInfo(@"C:\\Users\\Niteshkumar.sharma\\Desktop\\Demo\\Book.xlsx");
                    excelPackage.SaveAs(fi);
                }
            }
        }
    }
}
