using System;
using System.Windows;

namespace PraktikaWPF1YP
{
    public partial class AddOfferWindow : System.Windows.Window
    {
        public platezh NewOffer { get; set; }

        public AddOfferWindow() 
        {
            InitializeComponent();
            if (DpDate != null)
            {
                DpDate.SelectedDate = DateTime.Now;
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (TxtStatus == null || string.IsNullOrWhiteSpace(TxtStatus.Text))
            {
                MessageBox.Show("Укажите текущий статус отбора!", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (TxtSumma == null || !decimal.TryParse(TxtSumma.Text.Trim(), out decimal sum))
            {
                MessageBox.Show("Введите корректную сумму предлагаемого оклада 숫овым значением!", "Ошибка валидации", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            NewOffer = new platezh
            {
                data_platezha = DpDate?.SelectedDate ?? DateTime.Now,
                status = TxtStatus.Text.Trim(),
                summa = sum
            };

            DialogResult = true;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}




