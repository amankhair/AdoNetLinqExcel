using OfficeOpenXml;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Configuration;
using AdoNetLinq03.Tools;
using System.Data;

namespace AdoNetLinq03
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage package = new ExcelPackage();
            string pathForSaving = "\\Template";
            FileInfo template = new FileInfo("myTemplate.xlsx");
            
            using(package = new ExcelPackage(template, true))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                Helper helper = new Helper();
                DataSet data = helper.GetData();
                int row = 3;

                List<Area> areas = new List<Area>();

                foreach(DataRow dataRow in data.Tables["Area"].Rows)
                {
                    Area area = new Area();
                    area.AreaId = Int32.Parse(dataRow["AreaId"].ToString());
                    area.Name = dataRow["Name"].ToString();
                    area.FullName = dataRow["FullName"].ToString();
                    area.PavilionId = Int32.Parse(dataRow["PavilionId"].ToString());

                    areas.Add(area);
                }

                List<Pavilion> pavilions = new List<Pavilion>();
                foreach (DataRow dataRow in data.Tables["dic_Pavilion"].Rows)
                {
                    Pavilion pav = new Pavilion();
                    pav.PavilionId = Int32.Parse(dataRow["PavilionId"].ToString());
                    pav.Name = dataRow["Name"].ToString();
                    pavilions.Add(pav);
                }


                foreach (DataRow dataRow in data.Tables["Area"].Rows)
                {
                    worksheet.Cells[row, 2].Value = dataRow["Name"];
                    worksheet.Cells[row, 3].Value = dataRow["FullName"];
                    row++;
                }

                var query = areas.Join(pavilions,
                    a => a.PavilionId,
                    p => p.PavilionId,
                    (a, p) => new
                    {
                        a.Name,
                        a.FullName,
                        PavilionName = p.Name
                    });

                foreach (var dataRow in query)
                {
                    worksheet.Cells[row, 2].Value = dataRow.Name;
                    worksheet.Cells[row, 3].Value = dataRow.FullName;
                    worksheet.Cells[row, 7].Value = dataRow.PavilionName;

                    if (dataRow.PavilionName == "not determined")
                        Helper.SetColor(worksheet, "#AB0909", row, 2, row, 7);

                    row++;


                }

                

                package.SaveAs(new System.IO.FileInfo("demoOut.xlsx"));
            } 
        }
    }
}
