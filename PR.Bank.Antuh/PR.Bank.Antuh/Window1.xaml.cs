using System;
using System.Collections.Generic;
using System.Globalization;
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

namespace PR.Bank.Antuh
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        public Window1()
        {
            InitializeComponent();

        }

        private void bt_compare_Click(object sender, RoutedEventArgs e)
        {//передача данных
            var n = tbl_stab_result.Text;
            Window2 form = new Window2(n);
            form.Show();
        }

        private void sl_sum_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            NumberFormatInfo nfi = new NumberFormatInfo { NumberGroupSeparator = " ", NumberDecimalDigits = 0 };
            tb_sum.Text = ((Slider)sender).Value.ToString("n", nfi);
            try
            {
                double n = Convert.ToDouble(tb_sum.Text);
                double k = Convert.ToDouble(tb_srok.Text);
                double s = Convert.ToDouble(tb_popoln.Text);
                double m = n * k * s;
                tbl_stab_result.Text = Convert.ToString(m);
            }
            catch
            {
                MessageBox.Show("Error");
            }
        }

        private void sl_srok_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            NumberFormatInfo nfi = new NumberFormatInfo { NumberGroupSeparator = " ", NumberDecimalDigits = 0 };
            tb_srok.Text = ((Slider)sender).Value.ToString("n", nfi);

        }

        private void sl_popoln_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            NumberFormatInfo nfi = new NumberFormatInfo { NumberGroupSeparator = " ", NumberDecimalDigits = 0 };
            tb_popoln.Text = ((Slider)sender).Value.ToString("n", nfi);
        }
    }
}
