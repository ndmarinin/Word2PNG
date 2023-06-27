using System;
using System.IO;
using System.Windows;
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Win32;

namespace Word2JPG
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Load(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "(*.doc, *.docx)|*.doc;*.docx";
            if (openFileDialog.ShowDialog() == true)
                DocTextBox.Text = openFileDialog.FileName;
        }

        private void Button_Save(object sender, RoutedEventArgs e)
        {
            string path = new FileInfo(DocTextBox.Text).Directory.FullName;
            string file_name = new FileInfo(DocTextBox.Text).Name;


            var docPath = System.IO.Path.GetFullPath(DocTextBox.Text);
            var app = new Microsoft.Office.Interop.Word.Application();

            app.Visible = false;

            var doc = app.Documents.Open(docPath);

            doc.ShowGrammaticalErrors = false;
            doc.ShowRevisions = false;
            doc.ShowSpellingErrors = false;


            //Opens the word document and fetch each page and converts to image
            foreach (Microsoft.Office.Interop.Word.Window window in doc.Windows)
            {
                foreach (Microsoft.Office.Interop.Word.Pane pane in window.Panes)
                {
                    for (var i = 1; i <= pane.Pages.Count; i++)
                    {
                        var page = pane.Pages[i];
                        var bits = page.EnhMetaFileBits;
                        var target = System.IO.Path.Combine(path + "\\", string.Format("{1}_page_{0}.png", i, file_name.Split('.')[0]));

                        try
                        {
                            using (var ms = new MemoryStream((byte[])(bits)))
                            {
                                System.Drawing.Image image = System.Drawing.Image.FromStream(ms);
                                Bitmap bitmap = new Bitmap(image.Width, image.Height);

                                using (Graphics g = Graphics.FromImage(bitmap))
                                {
                                    //set background color
                                    g.Clear(System.Drawing.Color.White);
                                    using (SolidBrush brush = new SolidBrush(System.Drawing.Color.White))
                                    {
                                        g.FillRectangle(brush, 0, 0, bitmap.Width, bitmap.Height);
                                    }
                                    g.DrawImage(image, new System.Drawing.Rectangle(0, 0, image.Width, image.Height));
                                }

                                var pngTarget = System.IO.Path.ChangeExtension(target, "png");
                                bitmap.Save(pngTarget, ImageFormat.Png);
                            }
                        }
                        catch (System.Exception ex)
                        { }
                    }
                }
            }
            doc.Close(Type.Missing, Type.Missing, Type.Missing);
            app.Quit(Type.Missing, Type.Missing, Type.Missing);

            string messageBoxText = "Готово";
            string caption = "Word2PNG";
            MessageBoxButton button = MessageBoxButton.OK;

            MessageBox.Show(messageBoxText, caption, button);

        }
    }
}
