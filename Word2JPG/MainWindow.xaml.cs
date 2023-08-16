using System;
using System.IO;
using System.Windows;
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Win32;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using System.Linq.Expressions;

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
            openFileDialog.Filter = "(*.doc, *.docx, *.xls, *.xlsx)|*.doc;*.docx;*.xls;*.xlsx";
            if (openFileDialog.ShowDialog() == true)
                DocTextBox.Text = openFileDialog.FileName;
        }

        private void Button_Save(object sender, RoutedEventArgs e)
        {
            FileInfo fileInfo = new FileInfo(DocTextBox.Text);
            string path = fileInfo.Directory.FullName;
            string file_name = fileInfo.Name;


            string docPath = System.IO.Path.GetFullPath(DocTextBox.Text);
            if (fileInfo.Extension == ".xls" || fileInfo.Extension == ".xlsx")
            {
                //Start Excel and create a new document.
                Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
                oExcel.Visible = false;
                oExcel.Workbooks.Open(docPath);
                //var doc = app.ActiveSheet.Open(docPath);

                Microsoft.Office.Interop.Excel.Workbook wb = null;
                try
                {
                    wb = oExcel.Workbooks.Open(docPath.ToString(), false, false, Type.Missing, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", false, false, 0, false, true, 0);
                    Microsoft.Office.Interop.Excel.Sheets sheets = wb.Worksheets as Microsoft.Office.Interop.Excel.Sheets;
                    for (int j = 1; j <= sheets.Count; j++)
                    {
                        Microsoft.Office.Interop.Excel.Worksheet sheet = sheets[j];
                        //Following is used to find range with data
                        string startRange = "A4";
                        Microsoft.Office.Interop.Excel.Range endRange = sheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                        Microsoft.Office.Interop.Excel.Range range = sheet.get_Range(startRange, endRange);
                        range.Rows.AutoFit();
                        range.Columns.AutoFit();
                        range.Copy();

                        BitmapSource image = Clipboard.GetImage();
                        FormatConvertedBitmap fcbitmap = new FormatConvertedBitmap(image, PixelFormats.Bgr32, null, 0);
                        var target = System.IO.Path.Combine(path + "\\", string.Format("{1}_page_{0}.png", j, file_name.Split('.')[0]));
                        using (var fileStream = new FileStream(target, FileMode.Create))
                        {
                            PngBitmapEncoder encoder = new PngBitmapEncoder();
                            encoder.Interlace = PngInterlaceOption.On;
                            encoder.Frames.Add(BitmapFrame.Create(fcbitmap));
                            encoder.Save(fileStream);
                        }
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    wb.Close();
                    oExcel.Quit();
                    oExcel = null;
                }
            }
            else
            {
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
            }


            string messageBoxText = "Готово";
            Process.Start(@path);
            string caption = "Word2PNG";
            MessageBoxButton button = MessageBoxButton.OK;

            MessageBox.Show(messageBoxText, caption, button);

        }

        private void DocTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }
    }
}
