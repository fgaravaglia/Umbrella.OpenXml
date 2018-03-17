using System;

namespace Umbrella.Xlsx.DTOs
{
	/// <summary>
	/// DTO to fix limites of table to read
	/// </summary>
	public class DataTableDelimiter
	{
		/// <summary>
		/// Mode to terminate reading process
		/// </summary>
		public enum EnumEscapeMode
		{
			/// <summary>
			/// At first cell empty, stop reading
			/// </summary>
			FirstCellIsEmpty,
			/// <summary>
			/// stop reading when max number of rows is achieved
			/// </summary>
			MaximumRowNUmberAchieved
		}
		/// <summary>
		/// Index or row to start to read
		/// </summary>
		public int StartRowIndex { get; set; }
		/// <summary>
		/// Max column number ot read
		/// </summary>
		public int MaxColumnNumber { get; set; }
		/// <summary>
		/// Max row to read
		/// </summary>
		public int MaxRowNumber { get; private set; }
		/// <summary>
		/// HOw to terminate reading process
		/// </summary>
		public EnumEscapeMode RowEscapeMode { get; private set; }

		/// <summary>
		/// Default constructor
		/// </summary>
		public DataTableDelimiter()
		{
			StartRowIndex = 0;
			MaxColumnNumber = 99;
			RowEscapeMode = EnumEscapeMode.FirstCellIsEmpty;
			MaxRowNumber = 0;
		}
		/// <summary>
		/// Creates delimiter by mx row number achived
		/// </summary>
		/// <param name="maxRowNumber"></param>
		/// <returns></returns>
		public static DataTableDelimiter ByMaxRowNumber(int maxRowNumber)
		{
			if (maxRowNumber <= 0)
				throw new ArgumentOutOfRangeException(nameof(maxRowNumber), $"Max rownumber has to be greater than zero");
			var item = new DataTableDelimiter();
			item.RowEscapeMode = EnumEscapeMode.MaximumRowNUmberAchieved;
			item.MaxRowNumber = maxRowNumber;
			return item;
		}
	}
}
