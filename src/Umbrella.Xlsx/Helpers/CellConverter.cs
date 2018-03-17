using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Umbrella.Xlsx.DTOs;

namespace Umbrella.Xlsx.Helpers
{
	internal interface ICellConverter
	{
		Cell CreateNewCell(object value, ExcelColumnDescriptor.EnumColumnType cellType, string column = null, int rowNumber = 0);
	}

	internal class CellConverter : ICellConverter
	{
		readonly CultureInfo _CurrentCulture;
		readonly Dictionary<ExcelColumnDescriptor.EnumColumnType, Func<object, string>> _StringCellValueConverter;
		readonly Dictionary<ExcelColumnDescriptor.EnumColumnType, CellValues> _CellTypeConverter;
		readonly string _DateTimeFormat;

		public CellConverter(CultureInfo culture, string datetimeFormat)
		{
			this._CurrentCulture = culture ?? CultureInfo.InvariantCulture;
			this._DateTimeFormat = datetimeFormat ?? "dd/MM/yyyy";

			this._StringCellValueConverter = new Dictionary<ExcelColumnDescriptor.EnumColumnType, Func<object, string>>();
			this._CellTypeConverter = new Dictionary<ExcelColumnDescriptor.EnumColumnType, CellValues>();

			// add support
			AddSupportedCell(ExcelColumnDescriptor.EnumColumnType.String, val => val == null ? "" : val.ToString(), CellValues.String);
			AddSupportedCell(ExcelColumnDescriptor.EnumColumnType.Boolean, value => value == null ? "" : ((bool)value).ToString(), CellValues.String);
			AddSupportedCell(ExcelColumnDescriptor.EnumColumnType.Number,
							value =>
							{
								double? numbervalue = null;
								if (value != null)
									numbervalue = Convert.ToDouble(value, culture);
								return numbervalue.HasValue ? numbervalue.Value.ToString(culture) : "";
							}, CellValues.Number);
			AddSupportedCell(ExcelColumnDescriptor.EnumColumnType.Date,
							value =>
							{
								if (value == null)
									return "";

								// If I00m trying to store date as string, keep it as is. no parsing.
								if (value.GetType() == typeof(String))
									return (string)value;

								// try to parse it
								var dateValue = Convert.ToDateTime(value, culture);
								return dateValue.ToString(this._DateTimeFormat, culture);
							}, CellValues.Date);
			AddSupportedCell(ExcelColumnDescriptor.EnumColumnType.Formula, val => val == null ? "" : val.ToString(), CellValues.String);
		}

		public Cell CreateNewCell(object value, ExcelColumnDescriptor.EnumColumnType cellType, string column, int rowNumber)
		{
			if (String.IsNullOrEmpty(column))
				throw new ArgumentNullException(nameof(column));
			if (rowNumber == 0)
				throw new ArgumentOutOfRangeException(nameof(rowNumber), "The row number must to be greater or equals to 1");

			if (cellType == ExcelColumnDescriptor.EnumColumnType.Formula)
			{
				Cell cell = new Cell()
				{
					CellReference = $"{column}{rowNumber}"
				};
				CellFormula cellformula = new CellFormula(this._StringCellValueConverter[cellType].Invoke(value))
				{
					CalculateCell = true,

				};
				cell.Append(cellformula);
				return cell;
			}
			if (_StringCellValueConverter.ContainsKey(cellType))
			{
				return new Cell()
				{
					CellValue = new CellValue(this._StringCellValueConverter[cellType].Invoke(value)),
					DataType = new EnumValue<CellValues>(_CellTypeConverter[cellType])
				};
			}
			else
			{
				throw new NotImplementedException("unknown cell type " + cellType);
			}
		}

		private void AddSupportedCell(ExcelColumnDescriptor.EnumColumnType cellType, Func<object, string> converter, CellValues cellValueType)
		{
			if (_StringCellValueConverter.ContainsKey(cellType))
				throw new InvalidOperationException($"Unable to add {cellType}: item already registered");

			_StringCellValueConverter.Add(cellType, converter);
			_CellTypeConverter.Add(cellType, cellValueType);
		}
	}
}
