using System;
using System.Windows;

namespace PraktikaWPF1YP
{
    public partial class AddAgentWindow : System.Windows.Window
    {
        public agent NewAgent { get; private set; }

        public AddAgentWindow()
        {
            InitializeComponent();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtFam.Text) || string.IsNullOrWhiteSpace(TxtName.Text))
            {
                MessageBox.Show("Заполните обязательные поля: Фамилия и Имя сотрудника HR!", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            NewAgent = new agent
            {
                familiya = TxtFam.Text.Trim(),
                imya = TxtName.Text.Trim(),
                otchestvo = !string.IsNullOrWhiteSpace(TxtOtch.Text) ? TxtOtch.Text.Trim() : "HR",
                telefon = !string.IsNullOrWhiteSpace(TxtPhone.Text) ? TxtPhone.Text.Trim() : "Не указан"
            };

            DialogResult = true;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}
