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
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using GemBox.Spreadsheet;
using SpreadsheetLight;

namespace Test_app
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
            RunAsync();

        }

        static async Task RunAsync()
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                var workbook = new ExcelFile();
                var worksheet = workbook.Worksheets.Add("DATA");

                // Update port # in the following line.
                Uri geturi = new Uri("https://jsonplaceholder.typicode.com/comments"); //replace your url  
                System.Net.Http.HttpClient client = new System.Net.Http.HttpClient();
                System.Net.Http.HttpResponseMessage responseGet = await client.GetAsync(geturi);

                string response = await responseGet.Content.ReadAsStringAsync();

                //dynamic dynObj = JsonConvert.DeserializeObject(response);


                //string x = Convert.ToString(dynObj);



                List<ExcelList> lsObj = JsonConvert.DeserializeObject<List<ExcelList>>(response);

                int initial = 0;
               
                for (initial = 0; initial < 149; initial++)
                {
                    //System.Diagnostics.Debug.WriteLine("data " + initial.ToString());

                    worksheet.Cells[initial, 0].Value = lsObj[initial].id;
                    worksheet.Cells[initial, 1].Value = lsObj[initial].name;
                    worksheet.Cells[initial, 2].Value = lsObj[initial].email;
                }
                //worksheet.Cells.GetSubrangeAbsolute(4, 0, 4, 7).Merged = true;

                workbook.Save("DATA.xlsx");
           }
            catch (Exception e)
                {
                System.Diagnostics.Debug.WriteLine(e.ToString());
            }

        
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
        }
    }

    }
