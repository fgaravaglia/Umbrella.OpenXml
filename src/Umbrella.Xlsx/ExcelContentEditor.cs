using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Umbrella.Xlsx.DTOs;
using Umbrella.Xlsx.Helpers;

namespace Umbrella.Xlsx
{
	/// <summary>
	/// Class to edit excel 2007 content. Examples:
	/// <code>
	/// string folder = @"C:\Temp\Innova";
	///var table = new ExcelTableDescriptor()
	///{
	///	Name = "TestTable",
	///	SheetName = "Sheet01",
	///	Columns = new List<ExcelColumnDescriptor>()
	///						{
	///							new ExcelColumnDescriptor() { Name ="COL1", HeaderText="COl #1", ColumnType = ExcelColumnDescriptor.EnumColumnType.String },
	///							new ExcelColumnDescriptor() { Name ="COL2", HeaderText="COl #2", ColumnType = ExcelColumnDescriptor.EnumColumnType.Number },
	///							new ExcelColumnDescriptor() { Name ="COL3", HeaderText="COl #3", ColumnType = ExcelColumnDescriptor.EnumColumnType.Boolean },
	///							new ExcelColumnDescriptor() { Name ="COL4", HeaderText="COl #4", ColumnType = ExcelColumnDescriptor.EnumColumnType.String }
	///						}
	///};
	///
	///ExcelContentEditor editor = new ExcelContentEditor(CultureInfo.InvariantCulture)
	///		.SelectFolder(folder)
	///		.AddSheetTable(table)
	///		.FillTable("TestTable", new List<ExcelDataRowDescriptor>()
	///		{
	///						new ExcelDataRowDescriptor() {  RowNumber = 1 }.AddCell("COL1", "asjhg").AddCell("COL2", 22).AddCell("COL3", true).AddCell("COL4", DateTime.Now.ToString()),
	///						new ExcelDataRowDescriptor() {  RowNumber = 2 }.AddCell("COL1", "asjhg").AddCell("COL2", 22.5).AddCell("COL3", null).AddCell("COL4", null)
	///		})
	///		.SaveAs("Test.xlsx");
	///
	///
	///ExcelContentEditor reader = new ExcelContentEditor(CultureInfo.InvariantCulture);
	///var errors = reader.SelectFile(folder, "Test.xlsx")
	///			.OpenFile()
	///			.ReadTableFromSheet("Sheet01", new DataTableDelimiter())
	///			.ValidateTable(table);
	/// </code>
	/// To read an excel file instead:
	/// <code>
	/// string folder = @"C:\Temp\Innova";
	///var table = new ExcelTableDescriptor()
	///{
	///	Name = "WI",
	///	SheetName = "Work Items",
	///	Columns = new List<ExcelColumnDescriptor>()
	///						{
	///							new ExcelColumnDescriptor() { Name ="CR", HeaderText="CR", ColumnType = ExcelColumnDescriptor.EnumColumnType.String },
	///							new ExcelColumnDescriptor() { Name ="Title", HeaderText="Title", ColumnType = ExcelColumnDescriptor.EnumColumnType.String },
	///							new ExcelColumnDescriptor() { Name ="Customer", HeaderText="Customer", ColumnType = ExcelColumnDescriptor.EnumColumnType.String },
	///							new ExcelColumnDescriptor() { Name ="Priority", HeaderText="Priority", ColumnType = ExcelColumnDescriptor.EnumColumnType.String },
	///							new ExcelColumnDescriptor() { Name ="Status", HeaderText="Status", ColumnType = ExcelColumnDescriptor.EnumColumnType.String },
	///							new ExcelColumnDescriptor() { Name ="Assegnee", HeaderText="Assegnee", ColumnType = ExcelColumnDescriptor.EnumColumnType.String },
	///							new ExcelColumnDescriptor() { Name ="Effort", HeaderText="Effort [h]", ColumnType = ExcelColumnDescriptor.EnumColumnType.String },
	///							new ExcelColumnDescriptor() { Name ="ActualEffort", HeaderText="Actual Effort [h]", ColumnType = ExcelColumnDescriptor.EnumColumnType.String },
	///							new ExcelColumnDescriptor() { Name ="Delta", HeaderText="Delta Effort", ColumnType = ExcelColumnDescriptor.EnumColumnType.String }
	///						}
	///};
	///
	///ExcelContentEditor reader = new ExcelContentEditor(CultureInfo.InvariantCulture);
	///var errors = reader.SelectFile(folder, "TimeTrackingTemplate.xlsx")
	///			.OpenFile()
	///			.ReadTableFromSheet("Work Items", new DataTableDelimiter())
	///			.ValidateTable(table);
	///
	///		if (errors.Count == 0)
	///		{
	///			var dt = reader.GetTypedTable(table);
	///}
	/// 
	/// </code>
	/// 
	/// </summary>
	public class ExcelContentEditor
	{
		#region Fields

		string _TargetFolder;
		string _FileName;
		private byte[] _DocumentContent;
		Dictionary<string, int> _EmptySheets = new Dictionary<string, int>();
		Dictionary<string, ExcelTableDescriptor> _Tables = new Dictionary<string, ExcelTableDescriptor>();
		Dictionary<string, List<ExcelDataRowDescriptor>> _TableData = new Dictionary<string, List<ExcelDataRowDescriptor>>();
		Style _HeaderStyle;
		Style _RowStyle;
		DataTable _Data;
		readonly TableValidator _Validator;
		readonly ICellConverter _CellConverter;

		#endregion

		#region Properties

		public byte[] DocumentContent { get { return _DocumentContent; } }

		public DataTable ReadData { get { return _Data; } }

		#endregion

		internal ExcelContentEditor(ICellConverter converter)
		{
			if (converter == null)
				throw new ArgumentNullException(nameof(converter));

			this._CellConverter = converter;

		}
		public ExcelContentEditor(CultureInfo culture, string dateFormat) : this(new CellConverter(culture, dateFormat))
		{
			_HeaderStyle = new Style()
			{
				Bold = true,
				FontSize = 12,
				RbgColor = "FFFFFF"
			};

			_RowStyle = new Style() { FontSize = 10 };

			this._Validator = new TableValidator();
		}

		public ExcelContentEditor SelectFolder(string path)
		{
			this._TargetFolder = path;
			return this;
		}

		public ExcelContentEditor SelectFile(string path, string fileName)
		{
			_TargetFolder = path;
			_FileName = fileName;
			return this;
		}

		/// <summary>
		/// It adds an empty sheet
		/// </summary>
		/// <param name="name"></param>
		/// <returns></returns>
		public ExcelContentEditor AddEmptySheet(string name)
		{
			if (_EmptySheets.ContainsKey(name))
				throw new ArgumentException("Sheet already exists", nameof(name));

			var counter = _EmptySheets.Count;
			_EmptySheets.Add(name, counter + 1);

			return this;
		}
		/// <summary>
		/// It Adds the table to a sheet. Only one table per sheet is addmitted
		/// </summary>
		/// <param name="table"></param>
		/// <returns></returns>
		public ExcelContentEditor AddSheetTable(ExcelTableDescriptor table)
		{
			if (_Tables.ContainsKey(table.Name))
				throw new ArgumentException("Table already exists", nameof(table));
			if (_Tables.Count(x => x.Value.SheetName == table.SheetName) > 0)
				throw new ArgumentException($"Table already added to sheet {table.SheetName}", nameof(table));

			_Tables.Add(table.Name, table);

			return this;
		}
		/// <summary>
		/// Fills table with data
		/// </summary>
		/// <param name="tableName"></param>
		/// <param name="data"></param>
		/// <returns></returns>
		public ExcelContentEditor FillTable(string tableName, List<ExcelDataRowDescriptor> data)
		{

			if (_TableData.ContainsKey(tableName))
				throw new ArgumentException("Table already exists", nameof(tableName));

			_TableData.Add(tableName, data);

			return this;
		}
		/// <summary>
		/// Saves the output file with name
		/// </summary>
		/// <param name="fileName"></param>
		/// <returns></returns>
		public ExcelContentEditor SaveAs(string fileName)
		{
			Save();

			string outputFullpath = System.IO.Path.Combine(_TargetFolder, fileName);
			if (File.Exists(outputFullpath))
				File.Delete(outputFullpath);

			//Save the file with the new name
			File.WriteAllBytes(outputFullpath, _DocumentContent);

			return this;
		}
		/// <summary>
		/// Saves the output file in memory. Please get the conten from dedicated property
		/// </summary>
		public ExcelContentEditor Save()
		{
			using (MemoryStream stream = new MemoryStream())
			{
				//start editing
				using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
				{
					WorkbookPart workbookPart = document.AddWorkbookPart();
					workbookPart.Workbook = new Workbook();

					// add tables for sheets
					Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

					// Adding style
					WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
					stylePart.Stylesheet = GenerateStylesheet();
					stylePart.Stylesheet.Save();

					foreach (var tblDescriptor in this._Tables.Select(x => x.Value))
					{
						// create an empty sheet
						WorksheetPart worksheetPart = AddEmptySheet(workbookPart, tblDescriptor.SheetName, sheets);

						// Constructing header
						Row row = new Row();
						OpenXmlElement[] cells = tblDescriptor.Columns
												.Select(x => CreateTypedRowCell(1, x.HeaderText, ExcelColumnDescriptor.EnumColumnType.String, tblDescriptor.Columns.IndexOf(x), 2))
												.ToArray();
						row.Append(cells);
						// Insert the header row to the Sheet Data
						SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
						sheetData.AppendChild(row);
						worksheetPart.Worksheet.Save();

						// populate table
						var tblData = this._TableData[tblDescriptor.Name];
						int rowCounter = 1;
						foreach (var data in tblData)
						{
							rowCounter++;
							// create a new row
							Row currentRow = new Row();

							// build array of cell for this row
							OpenXmlElement[] rowCells = tblDescriptor.Columns.Select(x =>
							{
								if (!data.Values.ContainsKey(x.Name))
									throw new InvalidOperationException($"Unable to find a value for column {x.Name} inside table {tblDescriptor.Name}");

								// instance a new cell for proper type
								var cell = CreateTypedRowCell(rowCounter, data.Values[x.Name], x.ColumnType, tblDescriptor.Columns.IndexOf(x), 1);

								return cell;
							}).ToArray();

							currentRow.Append(rowCells);
							sheetData.AppendChild(currentRow);
						}
						worksheetPart.Worksheet.Save();
					}

					// add empty sheets
					foreach (var sheet in _EmptySheets.OrderBy(x => x.Value))
						document.AddSheet(sheet.Key);

					// if I'm using formula, caluclate when file is opened.
					if (workbookPart.Workbook.CalculationProperties != null)
					{
						workbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
						workbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
					}
					else
						workbookPart.Workbook.CalculationProperties = new CalculationProperties()
						{
							ForceFullCalculation = true,
							FullCalculationOnLoad = true
						};

					// save the file
					workbookPart.Workbook.Save();

				}
				//write on bytes
				_DocumentContent = stream.ToArray();
			}

			return this;
		}
		/// <summary>
		/// Reads in memory the file content
		/// </summary>
		/// <returns></returns>
		public ExcelContentEditor OpenFile()
		{
			string path = Path.Combine(this._TargetFolder, this._FileName);
			if (!File.Exists(path))
				throw new FileNotFoundException($"Unable to find specified file! pelase check target XLSX file", path);

			this._DocumentContent = File.ReadAllBytes(path);
			if (this.DocumentContent.Length == 0)
				throw new InvalidOperationException("File content is empty!");

			return this;
		}
		/// <summary>
		/// extracts the dataset from table stored inside sheet name
		/// </summary>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		public ExcelContentEditor ReadTableFromSheet(string sheetName, DataTableDelimiter delimiter)
		{
			if (this._Data != null)
				throw new InvalidOperationException($"Unable to read  file: override pre-existent data is forbidden!");
			if (String.IsNullOrEmpty(sheetName))
				throw new ArgumentNullException(nameof(sheetName));
			if (delimiter == null)
				throw new ArgumentNullException(nameof(delimiter));

			DataTable dt = new DataTable();
			using (MemoryStream stream = new MemoryStream())
			{
				// copy source file
				stream.Write(_DocumentContent, 0, (int)_DocumentContent.Length);
				//start editing
				using (SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true))
				{

					//Read the first Sheet from Excel file.
					Sheet sheet = doc.WorkbookPart.Workbook.Sheets.ChildElements.OfType<Sheet>().Single(x => x.Name == sheetName);

					//Get the Worksheet instance.
					Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;

					//Fetch all the rows present in the Worksheet.
					IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();

					//Loop through the Worksheet rows.
					foreach (Row row in rows)
					{
						int rowIndex = Array.IndexOf(rows.ToArray(), row);
						if (rowIndex < delimiter.StartRowIndex)
							continue;

						//Use the first row to add columns to DataTable.
						if (rowIndex == delimiter.StartRowIndex)
							CreateColumnSchema(row.Descendants<Cell>(), doc, dt, delimiter.MaxColumnNumber);
						else
						{
							//Add rows to DataTable.
							var cells = row.Descendants<Cell>();
							if (!CanAddThisRow(cells, delimiter, dt))
								break;

							dt.Rows.Add();
							int i = 0;
							foreach (Cell cell in cells)
							{
								// check on maximum side
								if (i == dt.Columns.Count)
									break;

								// fill the dataset
								var targetColumn = dt.Columns[i];
								dt.Rows[dt.Rows.Count - 1][i] = GetValue(doc, cell);
								i++;
							}
						}
					}
				}
			}

			this._Data = dt;
			return this;
		}
		/// <summary>
		/// extracts the dataset from table stored inside sheet name
		/// </summary>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		public List<string> ValidateTable(ExcelTableDescriptor tblDescriptor)
		{
			if (this._Data == null)
				throw new InvalidOperationException("Table is empty!!!");
			return this._Validator.ValidateTable(tblDescriptor, _Data);
		}
		/// <summary>
		/// Force column types of read data to table definition ones
		/// </summary>
		/// <param name="tblDescriptor"></param>
		/// <returns></returns>
		public DataTable GetTypedTable(ExcelTableDescriptor tblDescriptor)
		{
			var newTable = new DataTable();

			// create columns with name defined in table descriptor
			foreach (var col in tblDescriptor.Columns)
			{
				Type colType = null;
				switch (col.ColumnType)
				{
					case ExcelColumnDescriptor.EnumColumnType.Boolean:
						colType = typeof(bool);
						break;
					case ExcelColumnDescriptor.EnumColumnType.Number:
						colType = typeof(double);
						break;
					case ExcelColumnDescriptor.EnumColumnType.Date:
						colType = typeof(DateTime);
						break;
					case ExcelColumnDescriptor.EnumColumnType.String:
						colType = typeof(string);
						break;
				}
				newTable.Columns.Add(new DataColumn(col.Name, colType) { AllowDBNull = true });
			}

			//fill rows
			foreach (DataRow row in this._Data.Rows)
			{
				var newRow = newTable.NewRow();
				foreach (var col in tblDescriptor.Columns)
				{
					var index = tblDescriptor.Columns.IndexOf(col);
					if (col.ColumnType == ExcelColumnDescriptor.EnumColumnType.Date)
					{
						newRow[index] = DateTime.FromOADate(Convert.ToDouble(row[index]));
					}
					else
						newRow[index] = row[index];
				}
				newTable.Rows.Add(newRow);
			}

			return newTable;
		}

		#region Private Methods

		WorksheetPart AddEmptySheet(WorkbookPart workbookPart, string sheetName, Sheets sheets)
		{
			WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
			worksheetPart.Worksheet = new Worksheet();
			// craete new sheet ad append to file
			Sheet sheet = new Sheet()
			{
				Id = workbookPart.GetIdOfPart(worksheetPart),
				SheetId = sheets != null ? (UInt32)sheets.Count() + 1 : 1,
				Name = sheetName
			};
			sheets.Append(sheet);
			workbookPart.Workbook.Save();
			return worksheetPart;
		}

		private Cell CreateTypedRowCell(int rowNumber, object cellValue, ExcelColumnDescriptor.EnumColumnType colType, int columnIndexZeroBased, uint style = 1)
		{
			try
			{
				// get the name of column
				var columnName = GetColumnNameByIndex(columnIndexZeroBased);
				// craete cell
				var cell = this._CellConverter.CreateNewCell(cellValue, colType, columnName, rowNumber);
				// set proper style
				cell.StyleIndex = style;

				return cell;
			}
			catch (Exception ex)
			{
				throw new InvalidOperationException($"Unable to create cell of type {colType} in row {rowNumber}: {ex.Message} [Value: {cellValue}]", ex);
			}

		}

		private object GetValue(SpreadsheetDocument doc, Cell cell)
		{
			string value = cell?.CellValue?.InnerText;
			if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
			{
				return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
			}
			else if (cell.DataType != null && cell.DataType.Value == CellValues.Date)
			{
				return DateTime.FromOADate(Convert.ToDouble(value));
			}
			return value;
		}

		private Stylesheet GenerateStylesheet()
		{
			Stylesheet styleSheet = null;

			// create font for rows
			var rowFont = new Font()
			{
				FontSize = new FontSize() { Val = this._RowStyle.FontSize }
			};
			if (this._RowStyle.Bold)
				rowFont.Bold = new Bold() { Val = true };
			if (!String.IsNullOrEmpty(this._RowStyle.RbgColor))
				rowFont.Color = new Color() { Rgb = this._RowStyle.RbgColor };

			// create font for header
			var headerFont = new Font()
			{
				FontSize = new FontSize() { Val = this._HeaderStyle.FontSize }
			};
			if (this._HeaderStyle.Bold)
				headerFont.Bold = new Bold() { Val = true };
			if (!String.IsNullOrEmpty(this._HeaderStyle.RbgColor))
				headerFont.Color = new Color() { Rgb = this._HeaderStyle.RbgColor };

			Fonts fonts = new Fonts(
				rowFont,    //index 0 - rows
				headerFont // Index 1 - header
				);

			Fills fills = new Fills(
					new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
					new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
					new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })
					{ PatternType = PatternValues.Solid }) // Index 2 - header
				);

			Borders borders = new Borders(
					new Border(), // index 0 default
					new Border( // index 1 black border
						new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
						new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
						new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
						new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
						new DiagonalBorder())
				);

			CellFormats cellFormats = new CellFormats(
					new CellFormat(), // default
					new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
					new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true } // header
				);

			styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

			return styleSheet;
		}

		private void CreateColumnSchema(IEnumerable<Cell> cells, SpreadsheetDocument doc, DataTable dt, int maxColumnNumber)
		{
			int counter = 0;
			foreach (Cell cell in cells)
			{
				// limit the maximum number of read columns, to avoid memory exception
				if (counter == maxColumnNumber)
					return;

				var cellValue = GetValue(doc, cell);
				dt.Columns.Add(cellValue.ToString());
				counter++;
			}
		}

		private bool CanAddThisRow(IEnumerable<Cell> row, DataTableDelimiter delimiter, DataTable dt)
		{
			switch (delimiter.RowEscapeMode)
			{
				case DataTableDelimiter.EnumEscapeMode.FirstCellIsEmpty:
					var cell = (row.First().CellValue as CellValue)?.InnerText;
					return !String.IsNullOrEmpty(cell);

				case DataTableDelimiter.EnumEscapeMode.MaximumRowNUmberAchieved:
					return dt.Rows.Count < delimiter.MaxRowNumber;

				default:
					throw new NotImplementedException();
			}
		}

		private string GetColumnNameByIndex(int indexZeroBased)
		{
			string[] columnNames = new string[]
			{
				"A", "B", "C", "D", "E", "F", "G", "H", "I", "J","K", "L", "M", "N", "O","P", "Q", "R", "S", "T", "U", "V", "Z"
			};

			return columnNames[indexZeroBased];
		}
		#endregion
	}
}
