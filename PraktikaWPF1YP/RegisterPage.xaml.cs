using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace PraktikaWPF1YP
{
    public partial class RegisterPage : Page
    {
        public RegisterPage()
        {
            InitializeComponent();
        }

        private void BtnRegister_Click(object sender, RoutedEventArgs e)
        {
            // 1. Проверяем текстовые поля
            if (string.IsNullOrWhiteSpace(TxtLogin.Text) ||
                string.IsNullOrWhiteSpace(TxtPassword.Password) ||
                string.IsNullOrWhiteSpace(TxtConfirmPassword.Password))
            {
                MessageBox.Show("Внимание: Заполните ВСЕ поля формы!", "Проверка полей", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 2. Проверяем выпадающий список роли
            if (ComboRole.SelectedItem == null)
            {
                MessageBox.Show("Внимание: Вы не выбрали роль из списка!", "Проверка полей", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 3. Сверяем пароли
            if (TxtPassword.Password != TxtConfirmPassword.Password)
            {
                MessageBox.Show("Ошибка: Введенные пароли не совпадают!", "Проверка паролей", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string log = TxtLogin.Text.Trim();
            string hashedPassword = HashPassword(TxtPassword.Password);

            // 4. Получаем выбранный текст из ComboBox (русский текст)
            string comboBoxText = "";
            if (ComboRole.SelectedItem is ComboBoxItem item)
            {
                comboBoxText = item.Content.ToString();
            }
            else
            {
                comboBoxText = ComboRole.SelectedItem.ToString();
            }

            // 5. Конвертируем русскую роль в английский эквивалент, который требует база данных (CHECK constraint)
            string selectedRole = "User"; // Значение по умолчанию
            if (comboBoxText.Trim() == "Администратор")
            {
                selectedRole = "Admin";
            }
            else if (comboBoxText.Trim() == "Пользователь")
            {
                selectedRole = "User";
            }

            // Начинаем работу с базой данных
            try
            {
                using (YP_01_Praktic_1Entities db = new YP_01_Praktic_1Entities())
                {
                    // Проверяем существование пользователя
                    bool userExists = db.users.Any(u => u.login.ToLower() == log.ToLower());
                    if (userExists)
                    {
                        MessageBox.Show($"Ошибка: Логин '{log}' уже занят другим пользователем!", "Дубликат", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // Формируем объект с корректной английской ролью
                    users newUser = new users
                    {
                        login = log,
                        password_hash = hashedPassword,
                        role = selectedRole
                    };

                    db.users.Add(newUser);
                    db.SaveChanges(); // Теперь CHECK-ограничение будет пройдено успешно

                    MessageBox.Show("Регистрация успешно завершена!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);

                    if (NavigationService != null)
                    {
                        NavigationService.Navigate(new AuthPage());
                    }
                    else
                    {
                        MessageBox.Show("Критическая ошибка навигации: NavigationService равен null!", "Ошибка WPF", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (System.Data.Entity.Infrastructure.DbUpdateException dbEx)
            {
                var sqlMessage = dbEx.InnerException?.InnerException?.Message
                                 ?? dbEx.InnerException?.Message
                                 ?? dbEx.Message;

                MessageBox.Show($"[DbUpdateException] База отклонила запись!\n\nДетали от SQL сервера:\n{sqlMessage}",
                    "Ошибка уровня БД", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                StringBuilder completeError = new StringBuilder();
                completeError.AppendLine($"[Общая ошибка]: {ex.Message}");
                if (ex.InnerException != null)
                {
                    completeError.AppendLine($"[Внутреннее исключение]: {ex.InnerException.Message}");
                }

                MessageBox.Show(completeError.ToString(), "Критический сбой кода", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnGoToAuth_Click(object sender, RoutedEventArgs e)
        {
            if (NavigationService != null)
                NavigationService.Navigate(new AuthPage());
        }

        private string HashPassword(string password)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
                StringBuilder builder = new StringBuilder();
                foreach (var b in bytes)
                {
                    builder.Append(b.ToString("x2"));
                }
                return builder.ToString();
            }
        }
    }
}


