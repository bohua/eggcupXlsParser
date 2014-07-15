using System;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json;
using NDesk.Options;

namespace eggcupXlsParser
{
	class Program
	{
		static void Main(string[] args)
		{
			//bool argCheck = true;
			string tplPath = "";
			string expPath = "";
			string json = "";
			SheetModel inputs;

			OptionSet p = new OptionSet() {
				{
					"t|tpl=",
					"print template xls file.",
					v => tplPath = v
				},
				{
					"e|export=",
					"print xls file export path.",
					v => expPath = v
				},
				{
					"j|json=",
					"json string to map values into cells.",
					v => json = v
				}
			};

			try
			{
				p.Parse(args);
			}catch (Exception ex)
			{
				Console.Write("Error: PARSE_ARGS_ERROR::" + ex.Message);
				return;
			}

			try
			{
				inputs = JsonConvert.DeserializeObject<SheetModel>(json);
			}
			catch (Exception ex)
			{
				Console.Write("Error: PARSE_JSON_ERROR::" + ex.Message);
				return;
			}

			try{
				ISheet sheet;

				FileStream fileStream = new FileStream(tplPath, FileMode.Open);
				IWorkbook myWorkbook = new HSSFWorkbook(fileStream);

				sheet = myWorkbook.GetSheetAt(0);

				foreach (KeyValuePair<string, string> entry in inputs.singleMapper)
				{
					try
					{
						KeyValuePair<int, int> coordinate = getCoordinate(entry.Key);
						sheet.GetRow(coordinate.Value).GetCell(coordinate.Key).SetCellValue(entry.Value);
					}
					catch (Exception ex)
					{
						Console.Write("WARNING: PARSE_SINGLE_MAPPER_ERROR::" + ex.Message);
						continue;
					}
				}

				fileStream = new FileStream(expPath, FileMode.Create);
				myWorkbook.Write(fileStream);
				fileStream.Close();

			}
			catch (Exception ex)
			{
				Console.Write("Error: XLS_FILE_HANDLE_ERROR::" + ex.Message);
			}

			Console.ReadLine();

			return;
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
