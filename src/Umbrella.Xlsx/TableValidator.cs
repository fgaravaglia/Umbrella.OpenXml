using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using Umbrella.Xlsx.DTOs;

namespace Umbrella.Xlsx
{
	class TableValidator
	{
		readonly bool _ValidateColumnType;

		public TableValidator()
		{
			this._ValidateColumnType = false;
		}

		/// <summary>
		/// extracts the dataset from table stored inside sheet name
		/// </summary>
		/// <returns></returns>
		public List<string> ValidateTable(ExcelTableDescriptor tblDescriptor, DataTable data)
		{
			if (data == null)
				throw new ArgumentNullException(nameof(data));

			var errors = new List<string>();

			if (data.Columns.Count != tblDescriptor.Columns.Count)
				errors.Add($"Wrong Column number: expected <{tblDescriptor.Columns.Count}> but found <{data.Columns.Count}>");
			foreach (var col in tblDescriptor.Columns)
			{
				int columnIndex = tblDescriptor.Columns.IndexOf(col);
				if (data.Columns[columnIndex].ColumnName.Trim() != col.HeaderText.Trim())
					errors.Add($"Column {columnIndex}: Wrong Column Name; expected <{col.HeaderText}> but found <{data.Columns[columnIndex].ColumnName}>");
			}

			// check column types
			if (_ValidateColumnType)
			{
				foreach (var col in tblDescriptor.Columns)
				{
					int index = tblDescriptor.Columns.IndexOf(col);
					var columnType = data.Columns[index].DataType;
					switch (col.ColumnType)
					{
						case ExcelColumnDescriptor.EnumColumnType.String:
							if (columnType != typeof(string))
								errors.Add($"Column {col.Name}: expected {col.ColumnType} but found {columnType.Name}");
							break;
						case ExcelColumnDescriptor.EnumColumnType.Number:
							if (columnType != typeof(int) && columnType != typeof(int?)
								&& columnType != typeof(double) && columnType != typeof(double?))
								errors.Add($"Column {col.Name}: expected {col.ColumnType} but found {columnType.Name}");
							break;
						case ExcelColumnDescriptor.EnumColumnType.Date:
							if (columnType != typeof(DateTime?) && columnType != typeof(DateTime))
								errors.Add($"Column {col.Name}: expected {col.ColumnType} but found {columnType.Name}");
							break;
						case ExcelColumnDescriptor.EnumColumnType.Boolean:
							if (columnType != typeof(bool) && columnType != typeof(bool?))
								errors.Add($"Column {col.Name}: expected {col.ColumnType} but found {columnType.Name}");
							break;
						default:
							throw new NotImplementedException();
					}
				}
			}

			if (errors.Count > 0)
				return errors;

			//validate row data
			foreach (DataRow row in data.Rows)
			{
				StringBuilder error = new StringBuilder();
				foreach (var col in tblDescriptor.Columns)
				{
					object cellValue = row[col.HeaderText];
					if (cellValue == null)
						continue;
					switch (col.ColumnType)
					{
						case ExcelColumnDescriptor.EnumColumnType.String:
							break;
						case ExcelColumnDescriptor.EnumColumnType.Date:
							DateTime outDate;
							double oaDateValue;
							bool isOADate = Double.TryParse(cellValue.ToString(), out oaDateValue);

							if(!isOADate)
								error.AppendFormat("Cell {0}: invalid OA DateTime;", col.Name);
							else
							{
								if(!!String.IsNullOrEmpty(cellValue.ToString()) && !DateTime.TryParse(cellValue.ToString(), out outDate))
									error.AppendFormat("Cell {0}: invalid DateTime;", col.Name);
							}
							break;
						case ExcelColumnDescriptor.EnumColumnType.Number:
							double outNumber;
							if (!String.IsNullOrEmpty(cellValue.ToString()) && !double.TryParse(cellValue.ToString(), out outNumber))
								error.AppendFormat("Cell {0}: invalid number;", col.Name);
							break;
						case ExcelColumnDescriptor.EnumColumnType.Boolean:
							bool outBool;
							if (!String.IsNullOrEmpty(cellValue.ToString()) && !bool.TryParse(cellValue.ToString(), out outBool))
								error.AppendFormat("Cell {0}: invalid boolean;", col.Name);
							break;
						default:
							error.AppendFormat("Cell {0}: unknown format", col.Name);
							break;
					}

					if (error.Length > 0)
						errors.Add($"Row {data.Rows.IndexOf(row)} is Invali! " + error.ToString());
				}

			}
			return errors;
		}
	}
}
