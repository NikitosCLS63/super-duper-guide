using Microsoft.Win32;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Windows;
using Spire.Doc;

namespace WordLekcia
{
    public partial class Window3 : Window
    {
        private string selectedFilePath;

        public Window3()
        {
            InitializeComponent();
        }

        private void SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word files (*.docx)|*.docx|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                selectedFilePath = openFileDialog.FileName;
                MessageBox.Show("Выбран файл: " + selectedFilePath);
            }
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                MessageBox.Show("Пожалуйста, выберите файл для отправки.");
                return;
            }

            try
            {
               
                string tempFilePath = Path.GetTempFileName();

                
                Document doc = new Document();
                doc.LoadFromFile(selectedFilePath);

                doc.SaveToFile(tempFilePath, FileFormat.Docx);

                
                MailMessage messagev = new MailMessage(From.Text, To.Text, Subject.Text, "Пожалуйста, найдите прикрепленный документ.");

                Attachment attachment = new Attachment(tempFilePath);
                messagev.Attachments.Add(attachment);

                
                SmtpClient smtpClient = new SmtpClient("smtp.mail.ru", 587);
                messagev.Attachments.Add(new Attachment(selectedFilePath));
                string server = "smtp.mail.ru";
                string servergm = "smtp.mail.ru";
                SmtpClient smtpclient;

                if (server == "smtp.yandex.ru")
                {
                    smtpclient = new SmtpClient("smtp.yandex.ru", 587);
                }
                else if (servergm == "smtp.gmail.com")
                {
                    smtpclient = new SmtpClient("smtp.gmail.com", 587);
                }
                else if (server == "smtp.rambler.ru")
                {
                    smtpclient = new SmtpClient("smtp.rambler.ru", 465);
                }
                else
                {
                    smtpclient = new SmtpClient("smtp.mail.ru", 587);
                }
                try
                {
                    smtpclient.Send(messagev);
                    MessageBox.Show("Сообщение отправлено.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка отправки сообщения: " + ex.Message);
                }
                smtpclient.Credentials = new NetworkCredential(From.Text, Pass.Password);
                smtpclient.EnableSsl = true;
                smtpclient.Send(messagev);


                File.Delete(tempFilePath);

                MessageBox.Show("Сообщение отправлено успешно.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка отправки сообщения: " + ex.Message);
            }
        }
    }
}
