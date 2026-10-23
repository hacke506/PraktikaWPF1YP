using System;
using System.Windows;

namespace PraktikaWPF1YP
{
    public partial class AddCandidateWindow : System.Windows.Window
    {
        public klient NewCandidate { get; private set; }

        public AddCandidateWindow()
        {
            InitializeComponent();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            // Проверка обязательных полей (ФИО)
            if (string.IsNullOrWhiteSpace(TxtFam.Text) || string.IsNullOrWhiteSpace(TxtName.Text))
            {
                MessageBox.Show("Поля Фамилия и Имя обязательны к заполнению!", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Создание экземпляра модели клиента с учетом новых полей ввода
            NewCandidate = new klient
            {
                familiya = TxtFam.Text.Trim(),
                imya = TxtName.Text.Trim(),
                otchestvo = !string.IsNullOrWhiteSpace(TxtOtch.Text) ? TxtOtch.Text.Trim() : "Кандидатович",
                telefon = !string.IsNullOrWhiteSpace(TxtPhone.Text) ? TxtPhone.Text.Trim() : "89000000000",

                // Передача новых параметров в БД
                email = !string.IsNullOrWhiteSpace(TxtEmail.Text) ? TxtEmail.Text.Trim() : "Отсутствует",
                adres = !string.IsNullOrWhiteSpace(TxtAddress.Text) ? TxtAddress.Text.Trim() : "Не указано",

                // Системные значения по умолчанию для исключения ошибок NotNull ограничений в SQL Server
                data_rozhdeniya = DateTime.Now.AddYears(-22),
                pasport = "0000-000000"
            };

            DialogResult = true; // Закрытие окна с возвратом положительного флага
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}

