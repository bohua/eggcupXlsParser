using System;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json;
using NDesk.Options;
using NPOI.SS.Util;

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
			}
			catch (Exception ex)
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

			try
			{
				ISheet sheet;

				FileStream fileStream = new FileStream(tplPath, FileMode.Open);
				IWorkbook myWorkbook = new HSSFWorkbook(fileStream);

				sheet = myWorkbook.GetSheetAt(0);

				//Single Mapping Handling
				if (inputs.singleMapper != null)
				{
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
				}

				//Row Mapping Handling
				if (inputs.rowMapper != null)
				{
					int count = 0;
					foreach (IDictionary<string, string> row in inputs.rowMapper.rowData)
					{
						int destRowNum = inputs.rowMapper.startRow + count - 1;
						if (count > 0)
						{
							CopyRow((HSSFWorkbook)myWorkbook, (HSSFSheet)sheet, destRowNum - 1, destRowNum);
							//sheet.GetRow(distRowNum).HeightInPoints = sheet.GetRow(distRowNum - 1).HeightInPoints;
						}

						foreach (KeyValuePair<string, string> col in row)
						{
							KeyValuePair<int, int> coordinate = getCoordinate(col.Key + "," + (destRowNum + 1).ToString());
							sheet.GetRow(coordinate.Value).GetCell(coordinate.Key).SetCellValue(col.Value);
						}

						count++;
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

			//Console.ReadLine();

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

		#region LibFunc

		/// <summary>
		/// HSSFRow Copy Command
		///
		/// Description:  Inserts a existing row into a new row, will automatically push down
		///               any existing rows.  Copy is done cell by cell and supports, and the
		///               command tries to copy all properties available (style, merged cells, values, etc...)
		/// </summary>
		/// <param name="workbook">Workbook containing the worksheet that will be changed</param>
		/// <param name="worksheet">WorkSheet containing rows to be copied</param>
		/// <param name="sourceRowNum">Source Row Number</param>
		/// <param name="destinationRowNum">Destination Row Number</param>
		static public void CopyRow(HSSFWorkbook workbook, HSSFSheet worksheet, int sourceRowNum, int destinationRowNum)
		{
			// Get the source / new row
			IRow newRow = worksheet.GetRow(destinationRowNum);
			IRow sourceRow = worksheet.GetRow(sourceRowNum);

			// If the row exist in destination, push down all rows by 1 else create a new row
			if (newRow != null)
			{
				worksheet.ShiftRows(destinationRowNum, worksheet.LastRowNum, 1);
			}
			else
			{
				newRow = worksheet.CreateRow(destinationRowNum);
			}

			// Loop through source columns to add to new row
			for (int i = 0; i < sourceRow.LastCellNum; i++)
			{
				// Grab a copy of the old/new cell
				ICell oldCell = sourceRow.GetCell(i);
				ICell newCell = newRow.CreateCell(i);

				// If the old cell is null jump to next cell
				if (oldCell == null)
				{
					newCell = null;
					continue;
				}

				// Copy style from old cell and apply to new cell
				ICellStyle newCellStyle = workbook.CreateCellStyle();
				newCellStyle.CloneStyleFrom(oldCell.CellStyle); ;
				newCell.CellStyle = newCellStyle;

				// If there is a cell comment, copy
				if (newCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

				// If there is a cell hyperlink, copy
				if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;

				// Set the cell data type
				newCell.SetCellType(oldCell.CellType);

				// Set the cell data value
				switch (oldCell.CellType)
				{
					case CellType.Blank:
						newCell.SetCellValue(oldCell.StringCellValue);
						break;
					case CellType.Boolean:
						newCell.SetCellValue(oldCell.BooleanCellValue);
						break;
					case CellType.Error:
						newCell.SetCellErrorValue(oldCell.ErrorCellValue);
						break;
					case CellType.Formula:
						newCell.SetCellFormula(oldCell.CellFormula);
						break;
					case CellType.Numeric:
						newCell.SetCellValue(oldCell.NumericCellValue);
						break;
					case CellType.String:
						newCell.SetCellValue(oldCell.RichStringCellValue);
						break;
					case CellType.Unknown:
						newCell.SetCellValue(oldCell.StringCellValue);
						break;
				}
			}

			// If there are are any merged regions in the source row, copy to new row
			for (int i = 0; i < worksheet.NumMergedRegions; i++)
			{
				CellRangeAddress cellRangeAddress = worksheet.GetMergedRegion(i);
				if (cellRangeAddress.FirstRow == sourceRow.RowNum)
				{
					CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.RowNum,
																				(newRow.RowNum +
																				 (cellRangeAddress.FirstRow -
																				  cellRangeAddress.LastRow)),
																				cellRangeAddress.FirstColumn,
																				cellRangeAddress.LastColumn);
					worksheet.AddMergedRegion(newCellRangeAddress);
				}
			}


			worksheet.GetRow(destinationRowNum).HeightInPoints = worksheet.GetRow(sourceRowNum).HeightInPoints;
		}

		#endregion
	}
}
