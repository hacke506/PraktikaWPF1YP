using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using OfficeOpenXml;             // Для Excel (EPPlus)
using Xceed.Words.NET;           // Для Word (DocX)
using iTextSharp.text.pdf;       // Для PDF (iTextSharp)

namespace PraktikaWPF1YP
{
    public partial class DataViewPage : Page
    {
        private users activeUser; // Текущий залогиненный пользователь
        private bool isDarkTheme = false; // Флаг состояния текущей темы оформления

        public DataViewPage()
        {
            InitializeComponent();
            activeUser = new users { login = "HR_Admin_Test", role = "Administrator" };

            var mainWindow = Application.Current.MainWindow;
            if (mainWindow != null)
            {
                mainWindow.Title = "HR-Платформа JobMatch - Управление персоналом";
            }

            if (PanelAdminActions != null)
            {
                PanelAdminActions.Visibility = Visibility.Visible;
            }
        }

        public DataViewPage(users user)
        {
            InitializeComponent();
            activeUser = user;

            var mainWindow = Application.Current.MainWindow;
            if (mainWindow != null)
            {
                mainWindow.Title = $"JobMatch HR — Авторизован: {user.role}";
            }

            if (PanelAdminActions != null && (user.role == "Administrator" || user.role == "Администратор" || user.role == "SystemAdministrator"))
            {
                PanelAdminActions.Visibility = Visibility.Visible;
            }
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            RefreshData();
        }

        private void FilterChanged(object sender, SelectionChangedEventArgs e) { RefreshData(); }
        private void FilterChanged(object sender, TextChangedEventArgs e) { RefreshData(); }

        private void TablesTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshData();
        }

        private void RefreshData()
        {
            if (GridDogovors == null || GridKlients == null || GridAgents == null || GridPayments == null || TablesTabControl == null) return;

            string search = TxtSearch.Text.Trim().ToLower();
            int sortIndex = ComboSort.SelectedIndex;

            using (var db = new YP_01_Praktic_1Entities())
            {
                TabItem currentTab = TablesTabControl.SelectedItem as TabItem;
                if (currentTab == null) return;

                if (TablesTabControl.SelectedIndex == 0)
                {
                    var query = db.dogovor.AsQueryable();
                    if (!string.IsNullOrEmpty(search)) query = query.Where(d => d.nomer.ToLower().Contains(search));

                    var list = query.ToList();
                    if (sortIndex == 1) list = list.OrderBy(d => d.strahovaya_summa ?? 0).ToList();
                    if (sortIndex == 2) list = list.OrderByDescending(d => d.strahovaya_summa ?? 0).ToList();
                    GridDogovors.ItemsSource = list;
                }
                else if (TablesTabControl.SelectedIndex == 1)
                {
                    var query = db.klient.AsQueryable();
                    if (!string.IsNullOrEmpty(search)) query = query.Where(k => k.familiya.ToLower().Contains(search));

                    var list = query.ToList();
                    if (sortIndex == 1) list = list.OrderBy(k => k.id).ToList();
                    if (sortIndex == 2) list = list.OrderByDescending(k => k.id).ToList();
                    GridKlients.ItemsSource = list;
                }
                else if (TablesTabControl.SelectedIndex == 2)
                {
                    var query = db.agent.AsQueryable();
                    if (!string.IsNullOrEmpty(search)) query = query.Where(a => a.familiya.ToLower().Contains(search));

                    var list = query.ToList();
                    if (sortIndex == 1) list = list.OrderBy(a => a.id).ToList();
                    if (sortIndex == 2) list = list.OrderByDescending(a => a.id).ToList();
                    GridAgents.ItemsSource = list;
                }
                else if (TablesTabControl.SelectedIndex == 3)
                {
                    var query = db.platezh.AsQueryable();
                    if (!string.IsNullOrEmpty(search)) query = query.Where(p => p.status.ToLower().Contains(search));

                    var list = query.ToList();
                    if (sortIndex == 1) list = list.OrderBy(p => p.summa).ToList();
                    if (sortIndex == 2) list = list.OrderByDescending(p => p.summa).ToList();
                    GridPayments.ItemsSource = list;
                }
            }
        }

        private void NumberValidation(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !Regex.IsMatch(e.Text, "[0-9]");
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (var db = new YP_01_Praktic_1Entities())
                {
                    if (TablesTabControl.SelectedIndex == 0)
                    {
                        AddVacancyWindow win = new AddVacancyWindow();
                        win.Owner = Application.Current.MainWindow;
                        if (win.ShowDialog() == true && win.NewVacancy != null)
                        {
                            win.NewVacancy.nomer = $"POL-{DateTime.Now.Year}-{new Random().Next(100, 999)}";
                            db.dogovor.Add(win.NewVacancy);
                            db.SaveChanges();
                            MessageBox.Show("🎉 Новая вакансия успешно добавлена и сгенерирован шифр!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    else if (TablesTabControl.SelectedIndex == 1)
                    {
                        AddCandidateWindow win = new AddCandidateWindow();
                        win.Owner = Application.Current.MainWindow;
                        if (win.ShowDialog() == true && win.NewCandidate != null)
                        {
                            db.klient.Add(win.NewCandidate);
                            db.SaveChanges();
                            MessageBox.Show("🎉 Новый кандидат успешно внесен в базу данных!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    else if (TablesTabControl.SelectedIndex == 2)
                    {
                        AddAgentWindow win = new AddAgentWindow();
                        win.Owner = Application.Current.MainWindow;
                        if (win.ShowDialog() == true && win.NewAgent != null)
                        {
                            db.agent.Add(win.NewAgent);
                            db.SaveChanges();
                            MessageBox.Show("🎉 Сотрудник HR-агентства успешно добавлен!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    else if (TablesTabControl.SelectedIndex == 3)
                    {
                        AddOfferWindow win = new AddOfferWindow();
                        win.Owner = Application.Current.MainWindow;
                        if (win.ShowDialog() == true && win.NewOffer != null)
                        {
                            db.platezh.Add(win.NewOffer);
                            db.SaveChanges();
                            MessageBox.Show("🎉 Новый этап отбора соискателя зафиксирован!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                }
                RefreshData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Ошибка СУБД транзакции: {ex.Message}\n\nДетали: {ex.InnerException?.Message}", "Критический сбой", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show(
            "Вы действительно хотите удалить выбранную запись?\nВсе связанные с ней дочерние сущности будут уничтожены каскадно на SQL сервере!",
            "Подтверждение удаления", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes) return;
            using (var db = new YP_01_Praktic_1Entities())
            {
                if (TablesTabControl.SelectedIndex == 0 && GridDogovors.SelectedItem is dogovor d)
                    db.dogovor.Remove(db.dogovor.Find(Convert.ToInt32(d.id)));
                else if (TablesTabControl.SelectedIndex == 1 && GridKlients.SelectedItem is klient k)
                    db.klient.Remove(db.klient.Find(Convert.ToInt32(k.id)));
                else if (TablesTabControl.SelectedIndex == 2 && GridAgents.SelectedItem is agent a)
                    db.agent.Remove(db.agent.Find(Convert.ToInt32(a.id)));
                else if (TablesTabControl.SelectedIndex == 3 && GridPayments.SelectedItem is platezh p)
                    db.platezh.Remove(db.platezh.Find(Convert.ToInt32(p.id)));
                else
                {
                    MessageBox.Show("Выделите строку с записью в текущей таблице перед удалением!", "Внимание", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                db.SaveChanges();
            }
            RefreshData();
        }

        private void BtnShowChart_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                PraktikaWPF1YP.ChartWindow chartWin = new PraktikaWPF1YP.ChartWindow();
                chartWin.Owner = System.Windows.Application.Current.MainWindow;
                chartWin.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка вызова окна диаграммы: {ex.Message}");
            }
        }
        private void BtnToggleTheme_Click(object sender, RoutedEventArgs e)
        {
            isDarkTheme = !isDarkTheme;
            if (isDarkTheme)
            {
                MainGrid.Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));
                LblSearch.Foreground = Brushes.White;
                TxtSearch.Background = new SolidColorBrush(Color.FromRgb(45, 45, 48));
                TxtSearch.Foreground = Brushes.White;
                ComboSort.Background = new SolidColorBrush(Color.FromRgb(45, 45, 48));
                ComboSort.Foreground = Brushes.Black;
                var darkBg = new SolidColorBrush(Color.FromRgb(45, 45, 48));
                var darkRow = new SolidColorBrush(Color.FromRgb(37, 37, 38));
                var darkGridLines = new SolidColorBrush(Color.FromRgb(63, 63, 70));
                DataGrid[] grids = { GridDogovors, GridKlients, GridAgents, GridPayments };
                foreach (var grid in grids)
                {
                    if (grid == null) continue;
                    grid.Background = darkBg;
                    grid.RowBackground = darkRow;
                    grid.AlternatingRowBackground = darkBg;
                    grid.Foreground = Brushes.White;
                    grid.HorizontalGridLinesBrush = darkGridLines;
                    grid.VerticalGridLinesBrush = darkGridLines;
                }
                BtnToggleTheme.Content = "☀️ Светлая тема";
                BtnToggleTheme.Background = Brushes.Gold;
                BtnToggleTheme.Foreground = Brushes.Black;
            }
            else
            {
                MainGrid.Background = Brushes.Transparent;
                LblSearch.Foreground = Brushes.Black;
                TxtSearch.Background = Brushes.White;
                TxtSearch.Foreground = Brushes.Black;
                ComboSort.Background = Brushes.White;
                ComboSort.Foreground = Brushes.Black;
                var lightAlternateRow = new SolidColorBrush(Color.FromRgb(245, 247, 250));
                var lightGridLines = new SolidColorBrush(Color.FromRgb(220, 220, 220));
                DataGrid[] grids = { GridDogovors, GridKlients, GridAgents, GridPayments };
                foreach (var grid in grids)
                {
                    if (grid == null) continue;
                    grid.Background = Brushes.White;
                    grid.RowBackground = Brushes.White;
                    grid.AlternatingRowBackground = lightAlternateRow;
                    grid.Foreground = Brushes.Black;
                    grid.HorizontalGridLinesBrush = lightGridLines;
                    grid.VerticalGridLinesBrush = lightGridLines;
                }
                BtnToggleTheme.Content = "🌒 Тёмная тема";
                BtnToggleTheme.Background = new SolidColorBrush(Color.FromRgb(52, 73, 94));
                BtnToggleTheme.Foreground = Brushes.White;
            }
        }
        private string SafeGetProp(object obj, string name)
        {
            if (obj == null) return "";
            try
            {
                var p = obj.GetType().GetProperty(name);
                if (p != null)
                {
                    var val = p.GetValue(obj, null);
                    if (val is DateTime dt) return dt.ToString("dd.MM.yyyy");
                    return val?.ToString() ?? "";
                }
            }
            catch { }
            return "";
        }
        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Student");
            using (var db = new YP_01_Praktic_1Entities())
            {
                SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel Files (.xlsx)|.xlsx", FileName = $"Сводный_Отчет_HR_{DateTime.Now:yyyyMMdd}" };
                if (sfd.ShowDialog() == true)
                {
                    try
                    {
                        using (ExcelPackage package = new ExcelPackage())
                        {
                            ExcelWorksheet ws1 = package.Workbook.Worksheets.Add("Вакансии (Vacancies)");
                            ws1.Cells["A1"].Value = "Порядковый номер"; ws1.Cells["B1"].Value = "Шифр должности"; ws1.Cells["C1"].Value = "Дата публикации"; ws1.Cells["D1"].Value = "Дедлайн закрытия"; ws1.Cells["E1"].Value = "Предлагаемый оклад";
                            var dogovors = db.dogovor.ToList(); int row = 2;
                            foreach (var item in dogovors)
                            {
                                ws1.Cells[row, 1].Value = item.id; ws1.Cells[row, 2].Value = item.nomer; ws1.Cells[row, 3].Value = item.data_nachala?.ToString("dd.MM.yyyy"); ws1.Cells[row, 4].Value = item.data_okonchaniya?.ToString("dd.MM.yyyy"); ws1.Cells[row, 5].Value = item.strahovaya_summa; row++;
                            }
                            if (ws1.Dimension != null) ws1.Cells[ws1.Dimension.Address].AutoFitColumns();
                            ExcelWorksheet ws2 = package.Workbook.Worksheets.Add("Кандидаты (Candidates)");
                            ws2.Cells["A1"].Value = "Порядковый номер"; ws2.Cells["B1"].Value = "Фамилия соискателя"; ws2.Cells["C1"].Value = "Имя соискателя"; ws2.Cells["D1"].Value = "Отчество соискателя"; ws2.Cells["E1"].Value = "Контактный телефон";
                            var klients = db.klient.ToList(); row = 2;
                            foreach (var item in klients)
                            {
                                ws2.Cells[row, 1].Value = item.id; ws2.Cells[row, 2].Value = item.familiya; ws2.Cells[row, 3].Value = item.imya; ws2.Cells[row, 4].Value = item.otchestvo; ws2.Cells[row, 5].Value = item.telefon; row++;
                            }
                            if (ws2.Dimension != null) ws2.Cells[ws2.Dimension.Address].AutoFitColumns();
                            ExcelWorksheet ws3 = package.Workbook.Worksheets.Add("Рекрутеры (HR Agents)");
                            ws3.Cells["A1"].Value = "Порядковый номер"; ws3.Cells["B1"].Value = "Фамилия HR"; ws3.Cells["C1"].Value = "Имя HR"; ws3.Cells["D1"].Value = "Отчество HR"; ws3.Cells["E1"].Value = "Телефон";
                            var agents = db.agent.ToList(); row = 2;
                            foreach (var item in agents)
                            {
                                ws3.Cells[row, 1].Value = item.id; ws3.Cells[row, 2].Value = item.familiya; ws3.Cells[row, 3].Value = item.imya; ws3.Cells[row, 4].Value = item.otchestvo; ws3.Cells[row, 5].Value = item.telefon; row++;
                            }
                            if (ws3.Dimension != null) ws3.Cells[ws3.Dimension.Address].AutoFitColumns();
                            ExcelWorksheet ws4 = package.Workbook.Worksheets.Add("Офферы и Этапы (Offers)");
                            ws4.Cells["A1"].Value = "Порядковый номер"; ws4.Cells["B1"].Value = "Дата назначения этапа"; ws4.Cells["C1"].Value = "Текущий оклад"; ws4.Cells["D1"].Value = "Текущий статус";
                            var payments = db.platezh.ToList(); row = 2;
                            foreach (var item in payments)
                            {
                                ws4.Cells[row, 1].Value = item.id; ws4.Cells[row, 2].Value = item.data_platezha?.ToString("dd.MM.yyyy"); ws4.Cells[row, 3].Value = item.summa; ws4.Cells[row, 4].Value = item.status; row++;
                            }
                            if (ws4.Dimension != null) ws4.Cells[ws4.Dimension.Address].AutoFitColumns();
                            package.SaveAs(new FileInfo(sfd.FileName));
                        }
                        MessageBox.Show("🎉 Все 4 таблицы успешно выгружены в один файл Excel с раздельными листами!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex) { MessageBox.Show($"❌ Ошибка Excel: {ex.Message}"); }
                }
            }
        }
        // МЕТОД ИСПРАВЛЕН: Добавлен отсутствующий индекс [j] для Cells при записи текста параграфов Word
        private void ExportToWord_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new YP_01_Praktic_1Entities())
            {
                SaveFileDialog sfd = new SaveFileDialog { Filter = "Word Documents (.docx)|.docx", FileName = $"Сводный_Реестр_HR_{DateTime.Now:yyyyMMdd}" };
                if (sfd.ShowDialog() == true)
                {
                    try
                    {
                        using (DocX document = DocX.Create(sfd.FileName))
                        {
                            document.InsertParagraph("СВОДНЫЙ ОТЧЕТ ПЛАТФОРМЫ JOBMATCH").Bold().FontSize(16).Alignment = Xceed.Document.NET.Alignment.center;
                            document.InsertParagraph($"Дата формирования: {DateTime.Now:dd.MM.yyyy HH:mm}\n").FontSize(10);
                            // 1. Вакансии
                            document.InsertParagraph("1. Перечень открытых вакансий").Bold().FontSize(12);
                            var dogovors = db.dogovor.ToList();
                            var t1 = document.AddTable(dogovors.Count + 1, 5); t1.Design = Xceed.Document.NET.TableDesign.TableGrid;
                            t1.Rows[0].Cells[0].Paragraphs.First().Append("Порядковый номер");
                            t1.Rows[0].Cells[1].Paragraphs.First().Append("Шифр должности");
                            t1.Rows[0].Cells[2].Paragraphs.First().Append("Дата публикации");
                            t1.Rows[0].Cells[3].Paragraphs.First().Append("Дедлайн закрытия");
                            t1.Rows[0].Cells[4].Paragraphs.First().Append("Оклад (руб)");
                            for (int i = 0; i < dogovors.Count; i++)
                            {
                                t1.Rows[i + 1].Cells[0].Paragraphs.First().Append(dogovors[i].id.ToString());
                                t1.Rows[i + 1].Cells[1].Paragraphs.First().Append(dogovors[i].nomer ?? "");
                                t1.Rows[i + 1].Cells[2].Paragraphs.First().Append(dogovors[i].data_nachala?.ToString("dd.MM.yyyy") ?? "");
                                t1.Rows[i + 1].Cells[3].Paragraphs.First().Append(dogovors[i].data_okonchaniya?.ToString("dd.MM.yyyy") ?? "");
                                t1.Rows[i + 1].Cells[4].Paragraphs.First().Append(dogovors[i].strahovaya_summa?.ToString("F2") ?? "0.00");
                            }
                            document.InsertTable(t1); document.InsertParagraph("\n");
                            // 2. Кандидаты
                            document.InsertParagraph("2. База соискателей / кандидатов").Bold().FontSize(12);
                            var klients = db.klient.ToList();
                            var t2 = document.AddTable(klients.Count + 1, 5); t2.Design = Xceed.Document.NET.TableDesign.TableGrid;
                            t2.Rows[0].Cells[0].Paragraphs.First().Append("Порядковый номер");
                            t2.Rows[0].Cells[1].Paragraphs.First().Append("Фамилия соискателя");
                            t2.Rows[0].Cells[2].Paragraphs.First().Append("Имя соискателя");
                            t2.Rows[0].Cells[3].Paragraphs.First().Append("Контактный телефон");
                            t2.Rows[0].Cells[4].Paragraphs.First().Append("Дополнительно");
                            for (int i = 0; i < klients.Count; i++)
                            {
                                t2.Rows[i + 1].Cells[0].Paragraphs.First().Append(klients[i].id.ToString());
                                t2.Rows[i + 1].Cells[1].Paragraphs.First().Append(klients[i].familiya ?? "");
                                t2.Rows[i + 1].Cells[2].Paragraphs.First().Append(klients[i].imya ?? "");
                                t2.Rows[i + 1].Cells[3].Paragraphs.First().Append(klients[i].telefon ?? "");
                                t2.Rows[i + 1].Cells[4].Paragraphs.First().Append(SafeGetProp(klients[i], "pasport"));
                            }
                            document.InsertTable(t2); document.InsertParagraph("\n");
                            // 3. Рекрутеры
                            document.InsertParagraph("3. Список сотрудников HR-агентства").Bold().FontSize(12);
                            var agents = db.agent.ToList();
                            var t3 = document.AddTable(agents.Count + 1, 4); t3.Design = Xceed.Document.NET.TableDesign.TableGrid;
                            t3.Rows[0].Cells[0].Paragraphs.First().Append("Порядковый номер");
                            t3.Rows[0].Cells[1].Paragraphs.First().Append("Фамилия сотрудника");
                            t3.Rows[0].Cells[2].Paragraphs.First().Append("Имя сотрудника");
                            t3.Rows[0].Cells[3].Paragraphs.First().Append("Рабочий телефон");
                            for (int i = 0; i < agents.Count; i++)
                            {
                                t3.Rows[i + 1].Cells[0].Paragraphs.First().Append(agents[i].id.ToString());
                                t3.Rows[i + 1].Cells[1].Paragraphs.First().Append(agents[i].familiya ?? "");
                                t3.Rows[i + 1].Cells[2].Paragraphs.First().Append(agents[i].imya ?? "");
                                t3.Rows[i + 1].Cells[3].Paragraphs.First().Append(agents[i].telefon ?? "");
                            }
                            document.InsertTable(t3); document.InsertParagraph("\n");
                            // 4. Офферы
                            document.InsertParagraph("4. История офферов и этапы отбора").Bold().FontSize(12);
                            var payments = db.platezh.ToList();
                            var t4 = document.AddTable(payments.Count + 1, 4); t4.Design = Xceed.Document.NET.TableDesign.TableGrid;
                            t4.Rows[0].Cells[0].Paragraphs.First().Append("Порядковый номер");
                            t4.Rows[0].Cells[1].Paragraphs.First().Append("Дата этапа");
                            t4.Rows[0].Cells[2].Paragraphs.First().Append("Текущий статус");
                            t4.Rows[0].Cells[3].Paragraphs.First().Append("Оклад");
                            for (int i = 0; i < payments.Count; i++)
                            {
                                t4.Rows[i + 1].Cells[0].Paragraphs.First().Append(payments[i].id.ToString());
                                t4.Rows[i + 1].Cells[1].Paragraphs.First().Append(payments[i].data_platezha?.ToString("dd.MM.yyyy") ?? "");
                                t4.Rows[i + 1].Cells[2].Paragraphs.First().Append(payments[i].status ?? "");
                                t4.Rows[i + 1].Cells[3].Paragraphs.First().Append(payments[i].summa?.ToString("F2") ?? "0.00");
                            }
                            document.InsertTable(t4);
                            document.Save();
                        }
                        MessageBox.Show("🎉 Все данные успешно сохранены в документ Word в виде стандартных таблиц!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex) { MessageBox.Show($"❌ Ошибка Word: {ex.Message}"); }
                }
            }
        }
        private void ExportToPdf_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new YP_01_Praktic_1Entities())
            {
                SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF Files (.pdf)|.pdf", FileName = $"Сводный_Реестр_HR_{DateTime.Now:yyyyMMdd}" };
                if (sfd.ShowDialog() == true)
                {
                    try
                    {
                        iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 20, 20, 25, 25);
                        PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));
                        doc.Open();
                        string fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "Arial.ttf");
                        BaseFont bf = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                        iTextSharp.text.Font fontText = new iTextSharp.text.Font(bf, 10);
                        iTextSharp.text.Font fontTitle = new iTextSharp.text.Font(bf, 14, iTextSharp.text.Font.BOLD);
                        doc.Add(new iTextSharp.text.Paragraph("СВОДНЫЙ ОТЧЕТ ИЗ БАЗЫ ДАННЫХ JOBMATCH\n\n", fontTitle));
                        doc.Add(new iTextSharp.text.Paragraph("1. Перечень открытых вакансий компании", fontText));
                        PdfPTable table1 = new PdfPTable(5) { WidthPercentage = 100, SpacingBefore = 5f, SpacingAfter = 15f };
                        string[] h1 = { "Порядковый номер", "Шифр должности", "Дата публикации", "Дедлайн закрытия", "Оклад" };
                        foreach (var h in h1) table1.AddCell(new PdfPCell(new iTextSharp.text.Phrase(h, fontText)));
                        foreach (var item in db.dogovor.ToList())
                        {
                            table1.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.id.ToString(), fontText))); table1.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.nomer ?? "", fontText))); table1.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.data_nachala?.ToString("dd.MM.yyyy") ?? "", fontText))); table1.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.data_okonchaniya?.ToString("dd.MM.yyyy") ?? "", fontText))); table1.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.strahovaya_summa?.ToString("F2") ?? "0.00", fontText)));
                        }
                        doc.Add(table1);
                        doc.Add(new iTextSharp.text.Paragraph("2. База соискателей / кандидатов", fontText));
                        PdfPTable table2 = new PdfPTable(5) { WidthPercentage = 100, SpacingBefore = 5f, SpacingAfter = 15f };
                        string[] h2 = { "Порядковый номер", "Фамилия соискателя", "Имя соискателя", "Телефон", "Дополнительно" };
                        foreach (var h in h2) table2.AddCell(new PdfPCell(new iTextSharp.text.Phrase(h, fontText)));
                        foreach (var item in db.klient.ToList())
                        {
                            table2.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.id.ToString(), fontText))); table2.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.familiya ?? "", fontText))); table2.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.imya ?? "", fontText))); table2.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.telefon ?? "", fontText))); table2.AddCell(new PdfPCell(new iTextSharp.text.Phrase(SafeGetProp(item, "pasport"), fontText)));
                        }
                        doc.Add(table2);
                        doc.Add(new iTextSharp.text.Paragraph("3. Список сотрудников HR-агентства", fontText));
                        PdfPTable table3 = new PdfPTable(4) { WidthPercentage = 100, SpacingBefore = 5f, SpacingAfter = 15f };
                        string[] h3 = { "Порядковый номер", "Фамилия сотрудника", "Имя сотрудника", "Телефон" };
                        foreach (var h in h3) table3.AddCell(new PdfPCell(new iTextSharp.text.Phrase(h, fontText)));
                        foreach (var item in db.agent.ToList())
                        {
                            table3.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.id.ToString(), fontText))); table3.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.familiya ?? "", fontText))); table3.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.imya ?? "", fontText))); table3.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.telefon ?? "", fontText)));
                        }
                        doc.Add(table3);
                        doc.Add(new iTextSharp.text.Paragraph("4. История офферов и этапы отбора", fontText));
                        PdfPTable table4 = new PdfPTable(4) { WidthPercentage = 100, SpacingBefore = 5f, SpacingAfter = 5f };
                        string[] h4 = { "Порядковый номер", "Дата этапа", "Статус отбора", "Сумма" };
                        foreach (var h in h4) table4.AddCell(new PdfPCell(new iTextSharp.text.Phrase(h, fontText)));
                        foreach (var item in db.platezh.ToList())
                        {
                            table4.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.id.ToString(), fontText))); table4.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.data_platezha?.ToString("dd.MM.yyyy") ?? "", fontText))); table4.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.status ?? "", fontText))); table4.AddCell(new PdfPCell(new iTextSharp.text.Phrase(item.summa?.ToString("F2") ?? "0.00", fontText)));
                        }
                        doc.Add(table4);
                        doc.Close();
                        MessageBox.Show("🎉 Все данные успешно выгружены в один документ PDF со стандартными сетками!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex) { MessageBox.Show($"❌ Ошибка PDF: {ex.Message}"); }
                }
            }
        }
    }
}
