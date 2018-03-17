namespace Umbrella.Xlsx.DTOs
{
	public class ExcelColumnDescriptor
	{
		public enum EnumColumnType
		{
			String,
			Number,
			Date,
			Boolean,
			Formula
		}

		public string Name { get; set; }

		public string HeaderText { get; set; }

		public EnumColumnType ColumnType { get; set; }
	}
}
