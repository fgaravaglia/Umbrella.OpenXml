using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Umbrella.Xlsx.DTOs
{
	public class ExcelDataRowDescriptor
	{
		public int RowNumber { get; set; }

		public Dictionary<string, object> Values { get; set; }

		public ExcelDataRowDescriptor()
		{
			Values = new Dictionary<string, object>();
		}

		public ExcelDataRowDescriptor AddCell(string name, object cellValue)
		{
			this.Values.Add(name, cellValue);
			return this;
		}
	}
}
