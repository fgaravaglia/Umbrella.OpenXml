using System.Collections.Generic;

namespace Umbrella.Xlsx.DTOs
{
	/// <summary>
	/// Descriptor for table
	/// </summary>
	public class ExcelTableDescriptor
	{
		/// <summary>
		/// Name
		/// </summary>
		public string Name { get; set; }
		/// <summary>
		/// scheet that contains the table
		/// </summary>
		public string SheetName { get; set; }
		/// <summary>
		/// columns of table
		/// </summary>
		public List<ExcelColumnDescriptor> Columns { get; set; }
		/// <summary>
		/// Defualt Descritor
		/// </summary>
		public ExcelTableDescriptor()
		{
			this.Columns = new List<ExcelColumnDescriptor>();
		}
	}
}
