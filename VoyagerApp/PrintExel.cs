using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VoyagerApp
{
    class PrintExel
    {
        public static void ExportToExcel(List<InputData> Itog, string path )
        {
            // Загрузить Excel, затем создать новую пустую рабочую книгу
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // Сделать приложение Excel видимым
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel._Worksheet workSheet = excelApp.ActiveSheet;
            // Установить заголовки столбцов в ячейках
            workSheet.Cells[1, "A"] = "Номер счета";
            workSheet.Cells[1, "B"] = "Адрес";
            workSheet.Cells[1, "C"] = "Сумма долга";

           
            int row = 1;
            foreach (InputData c in Itog)
            {
                row++;
                workSheet.Cells[row, "A"] = c.LS;
                workSheet.Cells[row, "B"] = c.Address;
                workSheet.Cells[row, "C"] = c.DZ;
            }

            DateTime thisDay = DateTime.Today;
            string currentDate = thisDay.ToString("dd_MM_yy");
           
            excelApp.DisplayAlerts = false;
            workSheet.SaveAs(string.Format(@"{0}\Обход_" +currentDate+ ".xlsx", path));

            excelApp.Quit();

        }
    }
}
