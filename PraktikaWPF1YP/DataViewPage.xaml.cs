using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;


using OfficeOpenXml;             
using Xceed.Words.NET;           
using iTextSharp.text;           
using iTextSharp.text.pdf;       

namespace PraktikaWPF1YP
{
    public partial class DataViewPage : Page
    {
        public DataViewPage()
        {
            InitializeComponent();

           
            OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal("Student");

            LoadDataFromDatabase(); 
        }

        private void LoadDataFromDatabase()
        {
            try
            {
                using (var db = new YP_01_Praktic_1Entities())
                {
                    
                    var contractsList = db.dogovor.ToList();

                    
                    if (DataGridContracts != null)
                    {
                        DataGridContracts.ItemsSource = contractsList;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"⚠️ Ошибка при подключении к БД: {ex.Message}", "Ошибка загрузки", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string GetPremiumValue(object obj)
        {
            if (obj == null) return "";
            try
            {
                var properties = obj.GetType().GetProperties()
                    .Where(p => p.PropertyType.IsPrimitive ||
                                p.PropertyType == typeof(string) ||
                                p.PropertyType == typeof(decimal) ||
                                p.PropertyType == typeof(double) ||
                                p.PropertyType == typeof(int) ||
                                Nullable.GetUnderlyingType(p.PropertyType) != null)
                    .ToArray();

                if (properties.Length > 0)
                {
                    var val = properties[properties.Length - 1].GetValue(obj, null);
                    return val?.ToString() ?? "";
                }
            }
            catch { }
            return "";
        }

        // ================= ЭКСПОРТ В EXCEL (📊) =================
        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new YP_01_Praktic_1Entities())
            {
                var data = db.dogovor.ToList();
                if (!data.Any())
                {
                    MessageBox.Show("📭 Нет данных в базе для выполнения экспорта!", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = $"Отчет_Договоры_{DateTime.Now:yyyyMMdd}"
                };

                if (sfd.ShowDialog() == true)
                {
                    try
                    {
                        using (ExcelPackage package = new ExcelPackage())
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Договоры");

                            worksheet.Cells[1, 1].Value = "Номер полиса";
                            worksheet.Cells[1, 2].Value = "Дата начала";
                            worksheet.Cells[1, 3].Value = "Дата окончания";
                            worksheet.Cells[1, 4].Value = "Страховая сумма";
                            worksheet.Cells[1, 5].Value = "Премия";

                            for (int i = 1; i <= 5; i++)
                            {
                                worksheet.Cells[1, i].Style.Font.Bold = true;
                                worksheet.Cells[1, i].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            }

                            int row = 2;
                            foreach (var item in data)
                            {
                                worksheet.Cells[row, 1].Value = "";
                                worksheet.Cells[row, 2].Value = item.data_nachala?.ToString("dd.MM.yyyy") ?? "";
                                worksheet.Cells[row, 3].Value = item.data_okonchaniya?.ToString("dd.MM.yyyy") ?? "";
                                worksheet.Cells[row, 4].Value = item.strahovaya_summa;
                                worksheet.Cells[row, 5].Value = GetPremiumValue(item);
                                row++;
                            }

                            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                            File.WriteAllBytes(sfd.FileName, package.GetAsByteArray());
                        }
                        MessageBox.Show("🎉 Файл Excel успешно создан и сохранен!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"❌ Ошибка генерации Excel: {ex.Message}", "Сбой операции");
                    }
                }
            }
        }

        // ================= ЭКСПОРТ В WORD (📝) =================
        private void BtnExportWord_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new YP_01_Praktic_1Entities())
            {
                var data = db.dogovor.ToList();
                if (!data.Any()) { MessageBox.Show("📭 Данные отсутствуют!", "Внимание"); return; }

                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Word Documents (*.docx)|*.docx",
                    FileName = $"Отчет_Договоры_{DateTime.Now:yyyyMMdd}"
                };

                if (sfd.ShowDialog() == true)
                {
                    try
                    {
                        using (DocX document = DocX.Create(sfd.FileName))
                        {
                            var title = document.InsertParagraph("ОТЧЕТ ПО ДОГОВОРАМ СТРАХОВАНИЯ").Bold().FontSize(14);
                            title.Alignment = Xceed.Document.NET.Alignment.center;
                            document.InsertParagraph($"Сгенерировано автоматически: {DateTime.Now:dd.MM.yyyy HH:mm}\n\n");

                            var table = document.AddTable(data.Count + 1, 5);
                            table.Alignment = Xceed.Document.NET.Alignment.center;

                            string[] headers = { "Номер полиса", "Дата начала", "Дата окончания", "Страховая сумма", "Премия" };
                            for (int i = 0; i < headers.Length; i++)
                            {
                                var cell = table.Rows[0].Cells[i];
                                cell.Paragraphs.First().Append(headers[i]).Bold().Alignment = Xceed.Document.NET.Alignment.center;
                            }

                            int row = 1;
                            foreach (var item in data)
                            {
                                table.Rows[row].Cells[0].Paragraphs.First().Append("");
                                table.Rows[row].Cells[1].Paragraphs.First().Append(item.data_nachala?.ToString("dd.MM.yyyy") ?? "");
                                table.Rows[row].Cells[2].Paragraphs.First().Append(item.data_okonchaniya?.ToString("dd.MM.yyyy") ?? "");
                                table.Rows[row].Cells[3].Paragraphs.First().Append(item.strahovaya_summa?.ToString() ?? "0.00");
                                table.Rows[row].Cells[4].Paragraphs.First().Append(GetPremiumValue(item));
                                row++;
                            }

                            document.InsertTable(table);
                            document.Save();
                        }
                        MessageBox.Show("🎉 Файл Word успешно сохранен через проводник!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"❌ Ошибка генерации Word: {ex.Message}", "Сбой операции");
                    }
                }
            }
        }

        // ================= ЭКСПОРТ В PDF (📕) =================
        private void BtnExportPdf_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new YP_01_Praktic_1Entities())
            {
                var data = db.dogovor.ToList();
                if (!data.Any()) { MessageBox.Show("📭 Данные отсутствуют!", "Внимание"); return; }

                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    FileName = $"Отчет_Договоры_{DateTime.Now:yyyyMMdd}"
                };

                if (sfd.ShowDialog() == true)
                {
                    try
                    {
                        iTextSharp.text.Document doc = new iTextSharp.text.Document(PageSize.A4.Rotate(), 20, 20, 25, 25);
                        PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));
                        doc.Open();
                        string fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "Arial.ttf");
                        BaseFont bf = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                        iTextSharp.text.Font fontText = new iTextSharp.text.Font(bf, 10);
                        iTextSharp.text.Font fontHeader = new iTextSharp.text.Font(bf, 11, iTextSharp.text.Font.BOLD);
                        iTextSharp.text.Font fontTitle = new iTextSharp.text.Font(bf, 16, iTextSharp.text.Font.BOLD);
                        iTextSharp.text.Paragraph title = new iTextSharp.text.Paragraph("Выгрузка реестра договоров", fontTitle)
                        {
                            Alignment = iTextSharp.text.Element.ALIGN_CENTER,
                            SpacingAfter = 20
                        };
                        doc.Add(title);
                        PdfPTable table = new PdfPTable(5) { WidthPercentage = 100 };
                        table.SetWidths(new float[] { 2f, 2f, 2f, 2.5f, 2f });
                        string[] headers = { "Номер полиса", "Дата начала", "Дата окончания", "Страховая сумма", "Премия" };
                        foreach (string headerText in headers)
                        {
                            PdfPCell headerCell = new PdfPCell(new iTextSharp.text.Phrase(headerText, fontHeader))
                            {
                                HorizontalAlignment = iTextSharp.text.Element.ALIGN_CENTER,
                                VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE,
                                BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                                Padding = 6
                            };
                            table.AddCell(headerCell);
                        }
                        foreach (var item in data)
                        {
                            table.AddCell(CreateStyledCell("", fontText));
                            table.AddCell(CreateStyledCell(item.data_nachala?.ToString("dd.MM.yyyy"), fontText));
                            table.AddCell(CreateStyledCell(item.data_okonchaniya?.ToString("dd.MM.yyyy"), fontText));
                            table.AddCell(CreateStyledCell(item.strahovaya_summa?.ToString(), fontText));
                            table.AddCell(CreateStyledCell(GetPremiumValue(item), fontText));
                        }
                        doc.Add(table);
                        doc.Close();
                        MessageBox.Show("🎉 Данные успешно экспортированы один в один с таблицей!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"❌ Ошибка генерации PDF: {ex.Message}", "Сбой операции");
                    }
                }
            }
        }
        private PdfPCell CreateStyledCell(string text, iTextSharp.text.Font font)
        {
            return new PdfPCell(new iTextSharp.text.Phrase(text ?? "", font))
            {
                Padding = 5,
                VerticalAlignment = iTextSharp.text.Element.ALIGN_MIDDLE,
                HorizontalAlignment = iTextSharp.text.Element.ALIGN_LEFT
            };
        }
    }
}


