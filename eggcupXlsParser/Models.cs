using System;
using System.Collections.Generic;
using System.Text;

namespace eggcupXlsParser
{
	public class RowModel
	{
		public int startRow { get; set; }
		public IList<IDictionary<string, string>> rowData { get; set; }
	}

	public class SheetModel
	{
		public IDictionary<string, string> singleMapper { get; set; }
		public RowModel rowMapper { get; set; }
	}
}
