using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Umbrella.Xlsx.DTOs
{
	/// <summary>
	/// Descriptor for Data row
	/// </summary>
	public class ExcelDataRowDescriptor
	{
		/// <summary>
		/// Number of row
		/// </summary>
		public int RowNumber { get; set; }
		/// <summary>
		/// values of cells along the row
		/// </summary>
		public Dictionary<string, object> Values { get; set; }
		/// <summary>
		/// Default constructor
		/// </summary>
		public ExcelDataRowDescriptor()
		{
			Values = new Dictionary<string, object>();
		}
		/// <summary>
		/// Adds a new cell to row
		/// </summary>
		/// <param name="name"></param>
		/// <param name="cellValue"></param>
		/// <returns></returns>
		public ExcelDataRowDescriptor AddCell(string name, object cellValue)
		{
			this.Values.Add(name, cellValue);
			return this;
		}
	}
}
