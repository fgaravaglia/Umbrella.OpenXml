using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Umbrella.Xlsx.DTOs;

namespace Umbrella.Xlsx.Tests
{
	[TestClass]
	public class XlsxFileWriterTest : BaseTest
	{
		public override bool UseOutputFolder { get { return true; } }

		protected override void CleanUpTestData()
		{
			base.CleanUpTestData();
		}

		[TestMethod]
		public void WriteFileSucceeded()
		{
			//***** GIVEN
			string outputFileName = "XlsxFileWriterTest_WriteFileSucceeded.xlsx";
			var table = new ExcelTableDescriptor()
			{
				Name = "TestTable",
				SheetName = "Sheet01",
				Columns = new List<ExcelColumnDescriptor>()
				{
					 new ExcelColumnDescriptor() { Name = "COL1", HeaderText = "COl #1", ColumnType = ExcelColumnDescriptor.EnumColumnType.String },
					 new ExcelColumnDescriptor() { Name = "COL2", HeaderText = "COl #2", ColumnType = ExcelColumnDescriptor.EnumColumnType.Number },
					 new ExcelColumnDescriptor() { Name = "COL3", HeaderText = "COl #3", ColumnType = ExcelColumnDescriptor.EnumColumnType.Boolean },
					 new ExcelColumnDescriptor() { Name = "COL4", HeaderText = "COl #4", ColumnType = ExcelColumnDescriptor.EnumColumnType.String }
				}
			};
			var tableData = new List<ExcelDataRowDescriptor>()
			{
				 new ExcelDataRowDescriptor() { RowNumber = 1 }.AddCell("COL1", "asjhg").AddCell("COL2", 22).AddCell("COL3", true).AddCell("COL4", DateTime.Now.ToString()),
				 new ExcelDataRowDescriptor() { RowNumber = 2 }.AddCell("COL1", "asjhg").AddCell("COL2", 22.5).AddCell("COL3", null).AddCell("COL4", null)
			};

			//***** WHEN
			ExcelContentEditor editor = new ExcelContentEditor(CultureInfo.InvariantCulture, "dd/MM/yyyy")
				  .SelectFolder(this.OutputFolder)
				  .AddSheetTable(table)
				  .FillTable(table.Name, tableData)
				  .SaveAs(outputFileName);
			this._SavedFiles.Add(outputFileName);

			//***** ASSERT
			Assert.IsTrue(File.Exists(Path.Combine(OutputFolder, outputFileName)), $"Unable to find file {outputFileName} in folder {OutputFolder}");

			ExcelContentEditor reader = new ExcelContentEditor(CultureInfo.InvariantCulture, "dd/MM/yyyy")
					.SelectFile(this.OutputFolder, outputFileName)
					.OpenFile()
					.ReadTableFromSheet(table.SheetName, new DataTableDelimiter() { MaxColumnNumber = table.Columns.Count });
			Assert.IsNotNull(reader.ReadData as DataTable, "NUll data reade from file " + outputFileName);
			Assert.AreEqual(2, (reader.ReadData as DataTable).Rows.Count);
		}

		[TestMethod]
		public void WriteFileWithFormulaSucceeded()
		{
			//***** GIVEN
			string outputFileName = "XlsxFileWriterTest_WriteFileWithFormulaSucceeded.xlsx";
			var table = new ExcelTableDescriptor()
			{
				Name = "TestTable",
				SheetName = "Sheet01",
				Columns = new List<ExcelColumnDescriptor>()
				{
					 new ExcelColumnDescriptor() { Name = "COL1", HeaderText = "COl #1", ColumnType = ExcelColumnDescriptor.EnumColumnType.Number },
					 new ExcelColumnDescriptor() { Name = "COL2", HeaderText = "COl #2", ColumnType = ExcelColumnDescriptor.EnumColumnType.String },
					 new ExcelColumnDescriptor() { Name = "COL3", HeaderText = "COl #3", ColumnType = ExcelColumnDescriptor.EnumColumnType.Boolean },
					 new ExcelColumnDescriptor() { Name = "COL4", HeaderText = "COl #4", ColumnType = ExcelColumnDescriptor.EnumColumnType.String }
				}
			};
			var tableData = new List<ExcelDataRowDescriptor>()
			{
				 new ExcelDataRowDescriptor() { RowNumber = 1 }.AddCell("COL1", 1).AddCell("COL2", "title #1").AddCell("COL3", true).AddCell("COL4", DateTime.Now.ToString()),
				 new ExcelDataRowDescriptor() { RowNumber = 2 }.AddCell("COL1", 2).AddCell("COL2","title #2").AddCell("COL3", null).AddCell("COL4", null)
			};

			var table2 = new ExcelTableDescriptor()
			{
				Name = "FormulaTable",
				SheetName = "Sheet02",
				Columns = new List<ExcelColumnDescriptor>()
				{
					 new ExcelColumnDescriptor() { Name = "COL1", HeaderText = "ID", ColumnType = ExcelColumnDescriptor.EnumColumnType.Number },
					 new ExcelColumnDescriptor() { Name = "COL2", HeaderText = "ID2", ColumnType = ExcelColumnDescriptor.EnumColumnType.Number },
					 new ExcelColumnDescriptor() { Name = "COL3", HeaderText = "Title", ColumnType = ExcelColumnDescriptor.EnumColumnType.Formula },
				}
			};
			var fixedPartOfFormula = "VLOOKUP($A{0},'{1}'!$A$1:$Z$1000,{2})"; //=VLOOKUP(A2;Sheet01!A1:D5; 2)
			var tableData2 = new List<ExcelDataRowDescriptor>()
			{
				 new ExcelDataRowDescriptor() { RowNumber = 1 }
						.AddCell("COL1", 1)
						.AddCell("COL2", 4)
						.AddCell("COL3", String.Format(fixedPartOfFormula, 2, table.SheetName, 2)),
				 new ExcelDataRowDescriptor() { RowNumber = 1 }
						.AddCell("COL1", 1)
						.AddCell("COL2", 2)
						.AddCell("COL3", String.Format(fixedPartOfFormula, 3, table.SheetName, 2))
			};

			//***** WHEN
			ExcelContentEditor editor = new ExcelContentEditor(CultureInfo.InvariantCulture, "dd/MM/yyyy")
					.SelectFolder(this.OutputFolder)
					.AddSheetTable(table)
					.FillTable(table.Name, tableData)
					.AddSheetTable(table2)
					.FillTable(table2.Name, tableData2)
					.SaveAs(outputFileName);
			//this._SavedFiles.Add(outputFileName);

			//***** ASSERT
			Assert.IsTrue(File.Exists(Path.Combine(OutputFolder, outputFileName)), $"Unable to find file {outputFileName} in folder {OutputFolder}");
		}
	}
}
