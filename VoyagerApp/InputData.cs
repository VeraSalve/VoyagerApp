using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;

namespace VoyagerApp
{
    class InputData
    {
        public string LS { get; set; }
        public string Locality { get; set; }
        public string Street { get; set; }
        public string House { get; set; }
        public string Corps { get; set; }
        public string Flat { get; set; }
        public string Address { get; set; }
        public decimal DZ { get; set; }

        public InputData(string ls, string locality, string street, string house, string corps, string flat, string address, decimal dz)
        {
            this.LS = ls;
            this.Locality = locality;
            this.Street = street;
            this.House = house;
            this.Corps = corps;
            this.Flat = flat;
            this.Address = address;
            this.DZ = dz;
        }
        public static List<InputData> Parse(string path)
        {
            var result = new List<InputData>();

            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration() { FallbackEncoding = Encoding.GetEncoding(1252) }))
                {
                    var dataSet = reader.AsDataSet();
                    var table = dataSet.Tables[0];

                    for (int i = 1; i < table.Rows.Count; i++)
                    {
                        var ls = table.Rows[i][1].ToString();
                        var locality = table.Rows[i][5].ToString();
                        var street = table.Rows[i][8].ToString();

                        var house = table.Rows[i][9].ToString();
                        var corps = table.Rows[i][10].ToString();
                        var flat = table.Rows[i][11].ToString();
                        var tekDz = decimal.TryParse(table.Rows[i][24].ToString(), out decimal p) ? p : 0;
                        var dzZeroMonth = decimal.TryParse(table.Rows[i][27].ToString(), out decimal z) ? z : 0;
                        var address = locality + " " + street + " " + house + " " + corps;
                        var prdz = tekDz - dzZeroMonth;

                        result.Add(new InputData(ls, locality, street, house, corps, flat, address, prdz));
                    }
                }
            }

            return result;
        }





    }
}
