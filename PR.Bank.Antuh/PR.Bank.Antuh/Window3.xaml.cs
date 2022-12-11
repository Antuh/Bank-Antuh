using PR.Bank.Antuh.Module;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;


namespace PR.Bank.Antuh
{
    /// <summary>
    /// Логика взаимодействия для Window3.xaml
    /// </summary>
    public partial class Window3 : Window
    {
        private string name;
        private string kafedra;
        private string profession;
        private string groupe;
        public Window3(string names, string kafedras, string professions, string groupes)
        {
            InitializeComponent();
            name = names;
            kafedra = kafedras;
            profession = professions;
            groupe = groupes;
        }




        private readonly string TemplateFileName = @"D:\Download\Word.docx";//путь к файлу


        private void btn_voity_Click(object sender, RoutedEventArgs e)
        {
            string login = tb_login.Text;
            string password = tb_password.Password;
            Entities m = new Entities();
            var authorization = m.User;


            var user = authorization.Where(x => x.Login == login && x.Password == password).FirstOrDefault();
            if (user != null)
            {
                MessageBox.Show("Авторизация выполнена успешно");

                var wordApp = new Word.Application();//переменная для word
                wordApp.Visible = false;//word скрыт
                try
                {
                    var wordDocument = wordApp.Documents.Open(TemplateFileName);//переменная для хранения нашего документа

                    //Вставка вмето специальных выражений в нашем файле
                    ReplaceWordsStub("{name}", name, wordDocument);
                    ReplaceWordsStub("{kafedra}", kafedra, wordDocument);
                    ReplaceWordsStub("{profession}", profession, wordDocument);
                    ReplaceWordsStub("{groupe}", groupe, wordDocument);


                    wordDocument.SaveAs2(@"D:\Download\Word1.docx");//сохроняем наш документ
                    wordDocument.Close();//закрываем документ
                }
                catch
                {
                    MessageBox.Show("Произошла ошибка!!!");//окно ошибки
                }
                finally
                {
                    wordApp.Quit();//закрываем word
                }
            }
            else
            {
                MessageBox.Show("Кооректно введите логин и пароль");
            }


        }
        /// <summary>
        /// Метод замены ключевых слов на данные
        /// </summary>
        /// <param name="stubToReplace">Ключевые слова</param>
        /// <param name="text">Текст, который заменяет ключевые слова</param>
        /// <param name="wordDocument">Наш документ</param>
        private void ReplaceWordsStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;//перменная для хранения данных документа
            range.Find.ClearFormatting();//метод сброса всех натсроек текста
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);//находим ключевые слова и заменяем их
        }
    }
}
