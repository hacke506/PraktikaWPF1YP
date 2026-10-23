using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace PraktikaWPF1YP
{
    public partial class AuthPage : Page
    {
        private string currentCaptcha = "";

        public AuthPage()
        {
            InitializeComponent();
            GenerateCaptcha();
        }

        private void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtLogin.Text) || string.IsNullOrWhiteSpace(TxtPassword.Password))
            {
                MessageBox.Show("Заполните все поля!", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(TxtCaptchaInput.Text) || TxtCaptchaInput.Text.Trim().ToLower() != currentCaptcha.ToLower())
            {
                MessageBox.Show("Неверный код капчи!", "Ошибка капчи", MessageBoxButton.OK, MessageBoxImage.Error);
                GenerateCaptcha();
                return;
            }

            string hashedPassword = HashPassword(TxtPassword.Password);

            using (var db = new YP_01_Praktic_1Entities())
            {
                var user = db.users.FirstOrDefault(u => u.login == TxtLogin.Text && u.password_hash == hashedPassword);

                if (user != null)
                {
                    MessageBox.Show($"Добро пожаловать, {user.login}!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    NavigationService.Navigate(new DataViewPage());

                }
                else
                {
                    MessageBox.Show("Неверный логин или пароль!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    GenerateCaptcha();
                }
            }
        }

        private void BtnGoToRegister_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new RegisterPage());
        }

        private string HashPassword(string password)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
                StringBuilder builder = new StringBuilder();
                foreach (var b in bytes) builder.Append(b.ToString("x2"));
                return builder.ToString();
            }
        }

        private void GenerateCaptcha()
        {
            Random rand = new Random();
            string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            currentCaptcha = new string(Enumerable.Repeat(chars, 4).Select(s => s[rand.Next(s.Length)]).ToArray());
            TxtCaptchaCanvas.Text = currentCaptcha;
            TxtCaptchaInput.Text = "";
        }

        private void BtnRefreshCaptcha_Click(object sender, RoutedEventArgs e) => GenerateCaptcha();
    }
}


