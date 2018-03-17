using System;

namespace Umbrella.Xlsx.DTOs
{
	public class DataTableDelimiter
	{
		public enum EnumEscapeMode
		{
			FirstCellIsEmpty,
			MaximumRowNUmberAchieved
		}

		public int StartRowIndex { get; set; }

		public int MaxColumnNumber { get; set; }

		public int MaxRowNumber { get; private set; }

		public EnumEscapeMode RowEscapeMode { get; private set; }

		public DataTableDelimiter()
		{
			StartRowIndex = 0;
			MaxColumnNumber = 99;
			RowEscapeMode = EnumEscapeMode.FirstCellIsEmpty;
			MaxRowNumber = 0;
		}

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
