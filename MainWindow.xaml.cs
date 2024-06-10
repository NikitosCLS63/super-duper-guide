using Spire.Doc;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Xml.Linq;

namespace WordLekcia
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        //импорт
        //private void Button_Click(object sender, RoutedEventArgs e)
        //{
        //    Document doc = new Document();
        //    doc.LoadFromFile("ворд.docx");
        //    doc.SaveToFile("ртф.rtf", FileFormat.Rtf);

        //    TextRange range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
        //    FileStream fs = new FileStream("ртф.rtf", FileMode.OpenOrCreate);
        //    range.Load(fs, DataFormats.Rtf);
        //    fs.Close();





        //}
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Создаем объект OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Устанавливаем фильтр для открытия файлов формата docx
            openFileDialog.Filter = "Документы Word (*.docx)|*.docx";

            // Показываем диалоговое окно и проверяем, был ли выбран файл для открытия
            if (openFileDialog.ShowDialog() == true)
            {
                // Получаем путь к выбранному файлу
                string filePath = openFileDialog.FileName;

                // Создаем объект Document для загрузки документа Word
                Document document = new Document();
                document.LoadFromFile(filePath);

                // Извлекаем текст из документа Word
                string text = document.GetText();

                // Загружаем текст в RichTextBox
                rtb.Document.Blocks.Clear();
                rtb.Document.Blocks.Add(new Paragraph(new Run(text)));
            }
        }
        ////экспорт
        //private void Button_Click_1(object sender, RoutedEventArgs e)
        //{




        //    TextRange range = new TextRange(rtb.Document.ContentStart,rtb.Document.ContentEnd);
        //    FileStream fs = new FileStream("ртф.rtf", FileMode.Create);
        //    range.Save(fs, DataFormats.Rtf);
        //    fs.Close();

        //    Document doc = new Document();
        //    doc.LoadFromFile("ртф.rtf");

        //}


        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            // Создаем объект SaveFileDialog
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();

            // Устанавливаем фильтр для сохранения файлов в формате docx
            saveFileDialog.Filter = "Документы Word (*.docx)|*.docx";

            // Показываем диалоговое окно и проверяем, был ли выбран файл для сохранения
            if (saveFileDialog.ShowDialog() == true)
            {
                // Получаем путь к выбранному файлу
                string filePath = saveFileDialog.FileName;

                // Получаем содержимое RichTextBox
                TextRange range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);

                // Сохраняем содержимое RichTextBox в файл формата RTF
                string tempRtfFile = System.IO.Path.GetTempFileName();
                using (FileStream fs = new FileStream(tempRtfFile, FileMode.Create))
                {
                    range.Save(fs, DataFormats.Rtf);
                }

                // Создаем новый документ Word
                Document document = new Document();

                // Загружаем содержимое RTF из временного файла
                document.LoadFromFile(tempRtfFile, FileFormat.Rtf);

                // Сохраняем содержимое документа Word в файл формата DOCX по пути, указанному пользователем
                document.SaveToFile(filePath, FileFormat.Docx);

                // Удаляем временный файл RTF
                System.IO.File.Delete(tempRtfFile);
            }
        }









        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Window1 window = new Window1();
            window.Show();
            Close();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            Window3 window3 = new Window3();
            window3.Show();
            Close();
        }
    }
}