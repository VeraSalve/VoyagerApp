using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;

namespace VoyagerApp
{
    class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            
        }

      

       

        static class MyData
        {
            public static Microsoft.Office.Interop.Excel.Workbook xlWB1 { get; set; } //переменная будет отвечать за наш Excel файл
        }
    }
}
