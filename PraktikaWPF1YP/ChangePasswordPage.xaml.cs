using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Security.Cryptography;
using System.Text;

namespace PraktikaWPF1YP
{
    public partial class ChangePasswordPage : Page
    {
        private users account; // Профиль текущего пользователя

        public ChangePasswordPage(users user)
        {
            InitializeComponent();
            account = user;

            var mainWindow = Application.Current.MainWindow;
            if (mainWindow != null)
            {
                mainWindow.Title = "Смена пароля аккаунта";
            }
        }

        // Встроенный метод хэширования SHA-256 
        private string HashPassword(string password)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
                StringBuilder builder = new StringBuilder();
                foreach (byte b in bytes) builder.Append(b.ToString("x2"));
                return builder.ToString();
            }
        }

        private void BtnSavePassword_Click(object sender, RoutedEventArgs e)
        {
            if (TxtOldPass == null || TxtNewPass == null)
            {
                MessageBox.Show("Ошибка: Компоненты интерфейса не найдены в XAML!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string oldPass = TxtOldPass.Password.Trim();
            string newPass = TxtNewPass.Password.Trim();

            if (string.IsNullOrEmpty(oldPass) || string.IsNullOrEmpty(newPass))
            {
                MessageBox.Show("Заполните оба текстовых поля!", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string oldHash = HashPassword(oldPass);

            using (var db = new YP_01_Praktic_1Entities())
            {
                var userObj = db.users.FirstOrDefault(u => u.id == account.id && u.password_hash == oldHash);

                if (userObj != null)
                {
                    userObj.password_hash = HashPassword(newPass);
                    db.SaveChanges();

                    MessageBox.Show("Пароль успешно изменен в БД SQL Server!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    TxtOldPass.Password = "";
                    TxtNewPass.Password = "";
                }
                else
                {
                    MessageBox.Show("Старый пароль введен неверно!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}



