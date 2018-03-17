using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Umbrella.Xlsx.DTOs
{
	public class ExcelTableDescriptor
	{
		public string Name { get; set; }

		public string SheetName { get; set; }

		public List<ExcelColumnDescriptor> Columns { get; set; }

		public ExcelTableDescriptor()
		{
			this.Columns = new List<ExcelColumnDescriptor>();
		}
	}
}
