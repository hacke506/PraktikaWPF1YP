using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using Microsoft.Win32;
using OfficeOpenXml;             // Для работы с Excel (EPPlus)
using Xceed.Words.NET;           // Для работы с Word (DocX)
using Xceed.Document.NET;        // ОБЯЗАТЕЛЬНО: Исправляет ошибки Alignment и TableDesign

namespace PraktikaWPF1YP
{
    public partial class ChartWindow : global::System.Windows.Window
    {
        private bool isInitialized = false;

        public ChartWindow()
        {
            InitializeComponent();
            isInitialized = true;
        }

        private void ChartCanvas_Loaded(object sender, RoutedEventArgs e) { DrawChart(); }
        private void ChartCanvas_SizeChanged(object sender, SizeChangedEventArgs e) { DrawChart(); }
        private void FilterChanged(object sender, SelectionChangedEventArgs e) { DrawChart(); }
        private void FilterChanged(object sender, RoutedPropertyChangedEventArgs<double> e) { DrawChart(); }

        private void DrawChart()
        {
            if (!isInitialized || ChartCanvas == null) return;
            ChartCanvas.Children.Clear();

            try
            {
                using (var db = new YP_01_Praktic_1Entities())
                {
                    decimal minSalaryFilter = SliderMinSalary != null ? (decimal)SliderMinSalary.Value : 0;

                    var vacancies = db.dogovor
                                      .Where(d => d.strahovaya_summa >= minSalaryFilter)
                                      .OrderBy(d => d.strahovaya_summa)
                                      .Take(8)
                                      .ToList();

                    if (vacancies.Count == 0)
                    {
                        TextBlock emptyText = new TextBlock { Text = "Нет числовых данных для отображения по заданному фильтру!", FontSize = 13, Foreground = Brushes.Gray };
                        Canvas.SetLeft(emptyText, 150); Canvas.SetTop(emptyText, 140);
                        ChartCanvas.Children.Add(emptyText);
                        return;
                    }

                    double canvasWidth = ChartCanvas.ActualWidth;
                    double canvasHeight = ChartCanvas.ActualHeight;

                    decimal maxSalaryInDb = vacancies.Max(v => v.strahovaya_summa ?? 0);
                    double chartMaxLimit = Convert.ToDouble(maxSalaryInDb) * 1.25;
                    if (chartMaxLimit == 0) chartMaxLimit = 100000;

                    double stepX = canvasWidth / vacancies.Count;
                    double barWidth = stepX * 0.53;

                    int verticalSegments = 4;
                    for (int g = 0; g <= verticalSegments; g++)
                    {
                        double gridValue = (chartMaxLimit / verticalSegments) * g;
                        double gridY = canvasHeight - ((gridValue / chartMaxLimit) * (canvasHeight - 25));

                        Line gridLine = new Line { X1 = 0, Y1 = gridY, X2 = canvasWidth, Y2 = gridY, Stroke = new SolidColorBrush(Color.FromRgb(225, 228, 233)), StrokeThickness = 1 };
                        if (g == 0) gridLine.Stroke = Brushes.Black;
                        ChartCanvas.Children.Add(gridLine);

                        TextBlock lblY = new TextBlock { Text = $"{gridValue:N0}", FontSize = 9, Foreground = Brushes.DarkGray, TextAlignment = TextAlignment.Right, Width = 50 };
                        Canvas.SetLeft(lblY, -55); Canvas.SetTop(lblY, gridY - 7);
                        ChartCanvas.Children.Add(lblY);
                    }

                    Line axisY = new Line { X1 = 0, Y1 = 0, X2 = 0, Y2 = canvasHeight, Stroke = Brushes.Black, StrokeThickness = 1 };
                    ChartCanvas.Children.Add(axisY);

                    PointCollection polylinePoints = new PointCollection();

                    Color[] palette = {
                        Color.FromRgb(41, 128, 185), Color.FromRgb(39, 174, 96), Color.FromRgb(142, 68, 173),
                        Color.FromRgb(211, 84, 0),  Color.FromRgb(243, 156, 18), Color.FromRgb(44, 62, 80)
                    };

                    for (int i = 0; i < vacancies.Count; i++)
                    {
                        var item = vacancies[i];
                        double salaryValue = Convert.ToDouble(item.strahovaya_summa ?? 0);
                        double barHeight = (salaryValue / chartMaxLimit) * (canvasHeight - 25);

                        double rectLeft = (i * stepX) + (stepX - barWidth) / 2;
                        double rectTop = canvasHeight - barHeight;
                        double midPointX = (i * stepX) + (stepX / 2);

                        bool isLineChart = ComboChartType != null && ComboChartType.SelectedIndex == 1;

                        if (!isLineChart)
                        {
                            Rectangle rect = new Rectangle { Width = barWidth, Height = barHeight, Fill = new SolidColorBrush(palette[i % palette.Length]), RadiusX = 3, RadiusY = 3 };
                            Canvas.SetLeft(rect, rectLeft); Canvas.SetTop(rect, rectTop);
                            ChartCanvas.Children.Add(rect);
                        }
                        else
                        {
                            polylinePoints.Add(new Point(midPointX, rectTop));
                            Ellipse dot = new Ellipse { Width = 6, Height = 6, Fill = new SolidColorBrush(Color.FromRgb(231, 76, 60)) };
                            Canvas.SetLeft(dot, midPointX - 3); Canvas.SetTop(dot, rectTop - 3);
                            ChartCanvas.Children.Add(dot);
                        }

                        TextBlock txtValue = new TextBlock { Text = $"{salaryValue:N0} ₽", FontSize = 9, FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(Color.FromRgb(44, 62, 80)), TextAlignment = TextAlignment.Center, Width = stepX };
                        Canvas.SetLeft(txtValue, i * stepX); Canvas.SetTop(txtValue, rectTop - 16);
                        ChartCanvas.Children.Add(txtValue);

                        TextBlock txtLabel = new TextBlock { Text = item.nomer ?? $"ID {item.id}", FontSize = 9, FontWeight = FontWeights.SemiBold, Foreground = new SolidColorBrush(Color.FromRgb(120, 130, 140)), Width = stepX, TextAlignment = TextAlignment.Center };
                        Canvas.SetLeft(txtLabel, i * stepX); Canvas.SetTop(txtLabel, canvasHeight + 8);
                        ChartCanvas.Children.Add(txtLabel);
                    }

                    if (polylinePoints.Count > 1)
                    {
                        Polyline polyline = new Polyline { Points = polylinePoints, Stroke = new SolidColorBrush(Color.FromRgb(41, 128, 185)), StrokeThickness = 2 };
                        ChartCanvas.Children.Add(polyline);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка визуализации: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportChartToWord_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new YP_01_Praktic_1Entities())
            {
                decimal minSalaryFilter = SliderMinSalary != null ? (decimal)SliderMinSalary.Value : 0;
                var list = db.dogovor.Where(d => d.strahovaya_summa >= minSalaryFilter).ToList();

                SaveFileDialog sfd = new SaveFileDialog { Filter = "Документ Word (.docx)|*.docx", FileName = "Аналитика_Окладов_Вакансий" };
                if (sfd.ShowDialog() == true)
                {
                    try
                    {
                        using (DocX doc = DocX.Create(sfd.FileName))
                        {
                            doc.InsertParagraph("АНАЛИТИЧЕСКИЙ ОТЧЕТ ПРЕДЛАГАЕМЫХ ОКЛАДОВ").Bold().FontSize(14).Alignment = Alignment.center;
                            doc.InsertParagraph($"Порог фильтрации: от {minSalaryFilter:N0} ₽\nВыгружено вакансий: {list.Count}\nДата генерации: {DateTime.Now:dd.MM.yyyy HH:mm}\n").FontSize(10);

                            var table = doc.AddTable(list.Count + 1, 3);
                            table.Design = TableDesign.TableGrid;
                            table.Rows[0].Cells[0].Paragraphs.First().Append("ID Записи").Bold();
                            table.Rows[0].Cells[1].Paragraphs.First().Append("Шифр Должности").Bold();
                            table.Rows[0].Cells[2].Paragraphs.First().Append("Предлагаемый оклад").Bold();

                            for (int i = 0; i < list.Count; i++)
                            {
                                table.Rows[i + 1].Cells[0].Paragraphs.First().Append(list[i].id.ToString());
                                table.Rows[i + 1].Cells[1].Paragraphs.First().Append(list[i].nomer ?? "—");
                                table.Rows[i + 1].Cells[2].Paragraphs.First().Append($"{list[i].strahovaya_summa:N2} ₽");
                            }
                            doc.InsertTable(table);
                            doc.Save();
                        }
                        MessageBox.Show("🎉 Статистический отчет успешно выгружен в Word!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex) { MessageBox.Show($"Ошибка экспорта в Word: {ex.Message}"); }
                }
            }
        }

        private void ExportChartToExcel_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Student");
            using (var db = new YP_01_Praktic_1Entities())
            {
                decimal minSalaryFilter = SliderMinSalary != null ? (decimal)SliderMinSalary.Value : 0;
                var list = db.dogovor.Where(d => d.strahovaya_summa >= minSalaryFilter).ToList();
                SaveFileDialog sfd = new SaveFileDialog { Filter = "Книга Excel (.xlsx)|*.xlsx", FileName = "Аналитика_Окладов_Вакансий" };
                if (sfd.ShowDialog() == true)
                {
                    try
                    {
                        using (ExcelPackage pack = new ExcelPackage())
                        {
                            ExcelWorksheet ws = pack.Workbook.Worksheets.Add("Аналитика");
                            ws.Cells["A1"].Value = "Код"; ws.Cells["B1"].Value = "Шифр должности"; ws.Cells["C1"].Value = "Оклад соискателя (₽)";
                            ws.Row(1).Style.Font.Bold = true;
                            int row = 2;
                            foreach (var item in list)
                            {
                                ws.Cells[row, 1].Value = item.id;
                                ws.Cells[row, 2].Value = item.nomer;
                                ws.Cells[row, 3].Value = item.strahovaya_summa;
                                row++;
                            }
                            ws.Cells[ws.Dimension.Address].AutoFitColumns();
                            pack.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("🎉 Аналитические показатели успешно выгружены в таблицу Excel!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex) { MessageBox.Show($"Ошибка экспорта в Excel: {ex.Message}"); }
                }
            }
        }
    }
}