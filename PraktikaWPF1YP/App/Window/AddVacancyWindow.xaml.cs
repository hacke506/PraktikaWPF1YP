using System;
using System.Windows;

namespace PraktikaWPF1YP
{
    public partial class AddVacancyWindow : System.Windows.Window
    {
        public dogovor NewVacancy { get; private set; }

        public AddVacancyWindow()
        {
            InitializeComponent();
            DpStart.SelectedDate = DateTime.Now;
            DpEnd.SelectedDate = DateTime.Now.AddMonths(1);
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (!decimal.TryParse(TxtSalary.Text.Trim(), out decimal salary) || salary <= 0)
            {
                MessageBox.Show("Введите корректную положительную сумму предлагаемого оклада!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            NewVacancy = new dogovor
            {
                data_nachala = DpStart.SelectedDate ?? DateTime.Now,
                data_okonchaniya = DpEnd.SelectedDate ?? DateTime.Now.AddMonths(1),
                strahovaya_summa = salary
                // Поле nomer (шифр должности) заполняется автоматически в DataViewPage!
            };

            DialogResult = true;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}
