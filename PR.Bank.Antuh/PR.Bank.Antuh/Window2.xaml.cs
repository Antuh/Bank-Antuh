using System;
using System.Collections.Generic;
using System.IO;
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
    /// Логика взаимодействия для Window2.xaml
    /// </summary>
    public partial class Window2 : Window
    {
        public Window2()
        {
            InitializeComponent();
        }

        private void bt_vkladfour_Click(object sender, RoutedEventArgs e)
        {
            UIElement element = gd_screen as UIElement;
            Uri path = new Uri(@"D:\Download\screenshot.png");
            CaptureScreen(element, path);
        }
        public void CaptureScreen(UIElement source, Uri destination)
        {
            try
            {
                double Height, renderHeight, Width, renderWidth;

                Height = renderHeight = source.RenderSize.Height;
                Width = renderWidth = source.RenderSize.Width;


                RenderTargetBitmap renderTarget = new RenderTargetBitmap((int)renderWidth, (int)renderHeight, 96, 96, PixelFormats.Pbgra32);

                VisualBrush visualBrush = new VisualBrush(source);

                DrawingVisual drawingVisual = new DrawingVisual();
                using (DrawingContext drawingContext = drawingVisual.RenderOpen())
                {

                    drawingContext.DrawRectangle(visualBrush, null, new Rect(new Point(0, 0), new Point(Width, Height)));
                }

                renderTarget.Render(drawingVisual);


                PngBitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(renderTarget));
                using (FileStream stream = new FileStream(destination.LocalPath, FileMode.Create, FileAccess.Write))
                {
                    encoder.Save(stream);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
   
        private void bt_vkladthree_Click(object sender, RoutedEventArgs e)
        {
            var name = tbl_stability.Text;//фио студента
            var kafedra = tbl_stabilitydohod.Text;//кафедра
            var profession = tbl_stabilitystavka.Text;//специальность
            var groupe = tbl_stabilitysumma.Text;//группа
  
            Window3 form = new Window3(name, kafedra, profession, groupe);
            form.Show();
        }

        private void btn_vkladone_Click(object sender, RoutedEventArgs e)
        {
            /*var name = tbl_stability.Text;//фио студента
            var kafedra = tbl_stabilitydohod.Text;//кафедра
            var profession = tbl_stabilitystavka.Text;//специальность
            var groupe = tbl_stabilitysumma.Text;//группа
            Window3 form = new Window3(name, kafedra, profession, groupe);
            form.Show();*/
        }

        private void bt_vkladtwo_Click(object sender, RoutedEventArgs e)
        {
            /*var name = tbl_stability.Text;//фио студента
            var kafedra = tbl_stabilitydohod.Text;//кафедра
            var profession = tbl_stabilitystavka.Text;//специальность
            var groupe = tbl_stabilitysumma.Text;//группа
            Window3 form = new Window3(name, kafedra, profession, groupe);
            form.Show();*/
        }
    }
}
