using System;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace eggcupXlsParser
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
                        if (args.Length < 1)
                        {
                            Console.Write("Error: Please give at least one xls file path!");
                        }
                        else
                        {*/

            string json = @"{'B,3' : '测试公司名称' , 'E,3': '李小帅'}";

            Dictionary<string, string> inputs = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);

            ISheet sheet;
            try
            {
                FileStream fileStream = new FileStream(@"Templates\登记表.xls", FileMode.Open);
                //FileStream fileStream = new FileStream(args[0], FileMode.Open);
                IWorkbook myWorkbook = new HSSFWorkbook(fileStream);

                sheet = myWorkbook.GetSheetAt(0);

                foreach (KeyValuePair<string, string> entry in inputs)
                {
                    try
                    {
                        KeyValuePair<int, int> coordinate = getCoordinate(entry.Key);
                        sheet.GetRow(coordinate.Value).GetCell(coordinate.Key).SetCellValue(entry.Value);
                    }
                    catch (Exception ex) {
                        //Do nothing
                        continue;
                    }
                }

                fileStream = new FileStream(@"打印登记表.xls", FileMode.Create);
                myWorkbook.Write(fileStream);
                fileStream.Close();

            }
            catch (Exception ex)
            {
                Console.Write("Error:" + ex.Message);
            }
            //}
            Console.ReadLine();
        }

        static readonly string[] Columns = new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH" };
        static KeyValuePair<int, int> getCoordinate(string cellName)
        {
            string[] coordinates = cellName.Split(',');
            int x = Array.IndexOf(Columns, coordinates[0]);
            int y;
            bool parsed = Int32.TryParse(coordinates[1], out y);

            if (x < 0 || !parsed)
            {
                throw new Exception("NOT_VALID_CELL_NAME");
            }

            return new KeyValuePair<int, int>(x, y - 1);
        }
    }
}
