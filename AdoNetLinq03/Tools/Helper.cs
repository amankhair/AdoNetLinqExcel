using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AdoNetLinq03.Tools
{
    class Helper
    {
        public DataSet GetData()
        {
            string conString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            using (SqlConnection con = new SqlConnection(conString))
            {
                con.Open();

                SqlCommand cmd = new SqlCommand();
                
                cmd.Connection = con;
                cmd.CommandText = "SELECT * FROM [CRCMS_new].[dbo].[Area];" +
                                    "SELECT* FROM[CRCMS_new].[dbo].[dic_Group];" +
                                    "SELECT* FROM[CRCMS_new].[dbo].[dic_Pavilion];";

                DataSet ds = new DataSet();
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.TableMappings.Add("Table", "Area");
                da.TableMappings.Add("Table1", "dic_Group");
                da.TableMappings.Add("Table2", "dic_Pavilion");

                da.Fill(ds);
                con.Dispose();
                return ds;
            }
        }

        private void SetColor(ExcelWorksheet worksheet, string color, int FRow, int FCol, int ToRow, int ToCol)
        {
            using (ExcelRange rng = worksheet.Cells[FRow, FCol, ToRow, ToCol])
            {
                if (color != null)
                {
                    System.Drawing.Color colContractHex =
                        ColorTranslator.FromHtml(color);

                    rng.Style.Fill.PatternType =
                        ExcelFillStyle.Solid;

                    rng.Style.Fill.BackgroundColor
                        .SetColor(colContractHex);
                }

                //rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                //rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                //rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                //rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                //rng.Style.WrapText = true;
            }
        }

    }

    
    public class Area
    {
        public int AreaId { get; set; }
        public string Name { get; set; }
        public string FullName { get; set; }
        public int PavilionId { get; set; }
    }

    public class Pavilion
    {
        public int PavilionId { get; set; }
        public string Name { get; set; }
    }
}
