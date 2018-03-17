namespace Umbrella.Xlsx.DTOs
{
	/// <summary>
	/// Descriptor for column on excel
	/// </summary>
	public class ExcelColumnDescriptor
	{
		/// <summary>
		/// SUpported types of column
		/// </summary>
		public enum EnumColumnType
		{
			/// <summary>
			/// Text column
			/// </summary>
			String,
			/// <summary>
			/// Number
			/// </summary>
			Number,
			/// <summary>
			/// Date
			/// </summary>
			Date,
			/// <summary>
			/// boolean
			/// </summary>
			Boolean,
			/// <summary>
			/// Formula
			/// </summary>
			Formula
		}
		/// <summary>
		/// COlumn name
		/// </summary>
		public string Name { get; set; }
		/// <summary>
		/// Header text
		/// </summary>
		public string HeaderText { get; set; }
		/// <summary>
		/// Type of column
		/// </summary>
		public EnumColumnType ColumnType { get; set; }
	}
}
