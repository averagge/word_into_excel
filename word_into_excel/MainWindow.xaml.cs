using Microsoft.Win32;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Threading;

namespace word_into_excel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private WordClass word;


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.doc;*.docx",
                Filter = "Word Documents (*.docx, *.doc)|*.docx;*.doc",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            wordlist.Items.Add(ofd.FileName);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx;*.xls",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            excelfile.Text = ofd.FileName;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            try
            {
                word = new WordClass();
                foreach (var item in wordlist.Items)
                {
                    word.GetAllData(item.ToString());
                }
                word.WriteIntoExcel(excelfile.Text);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }


        }
        private void wordlist_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

            if (sender is ListBox listBox)
            {
                if (listBox.SelectedItem != null)
                {
                    listBox.Items.Remove(listBox.SelectedItem);
                }
            }

        }
    }
}
