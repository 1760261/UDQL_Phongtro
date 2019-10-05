using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace _1760261_1760459
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

        private void Import_Button_Click(object sender, RoutedEventArgs e)
        {
            var Screen = new OpenFileDialog();
          
                if (Screen.ShowDialog() == true)
                {
                if (System.IO.Path.GetExtension(Screen.FileName) != ".xlsx")
                {
                    MessageBox.Show("Please chose excel file");
                    return;
                }
                else
                {
                 
                    var workbook = new Aspose.Cells.Workbook(Screen.FileName);
                    var col = "A";
                    var row = 2;
                    var sheet = workbook.Worksheets[0];
                    var cell = sheet.Cells[$"{col}{row}"];
                    var db = new PhongTroEntities();
                    while (cell.Value != null)
                    {
                      
                        var Price = sheet.Cells[$"A{row}"].IntValue;
                        var Type = sheet.Cells[$"D{row}"].StringValue;
                        var Status = sheet.Cells[$"C{row}"].BoolValue;
                        var ImagePath = sheet.Cells[$"E{row}"].StringValue;
                        var Description = sheet.Cells[$"B{row}"].StringValue;

                        var sourceImage = new FileInfo(Screen.FileName);
                        Debug.WriteLine(sourceImage.DirectoryName);
                        var fullImageSource = $"{ sourceImage.DirectoryName}\\{ ImagePath}";
                        var sourceImageInfo = new FileInfo(fullImageSource);
                        var uniqueName =$"{ Guid.NewGuid()}{sourceImageInfo.Extension}";
                        var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                        var destinationBase = $"{baseDirectory}\\import\\{uniqueName}";

                        File.Copy(fullImageSource, destinationBase);
                        var newRoom = new Room()
                        {
                            price = Price,
                            type = Type,
                            status = Status,
                            imagePath = uniqueName,
                            description = Description
                        };
                        db.Rooms.Add(newRoom);
                        db.SaveChanges();
                        row++;
                        cell = sheet.Cells[$"{col}{row}"];

                    }
                }
                var db2 = new PhongTroEntities();
                RoomListView.ItemsSource = db2.Rooms.ToList();
                MessageBox.Show("Imported");
                }
            
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var db = new PhongTroEntities();
            RoomListView.ItemsSource = db.Rooms.ToList();
        }
    }
}
