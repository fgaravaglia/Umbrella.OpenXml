using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Umbrella.Docx.Helpers;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace Umbrella.Docx
{
	/// <summary>
	/// Class to edit docx file
	/// </summary>
	///// <example>
	///// string folder = @"C:\Temp\Innova";
	///// string template = "TemplateFattura.docx";
	///// DocxContentEditor editor = new DocxContentEditor().SelectFile(folder, template).OpenFile()
	///// 	.UpdateProperty("_DISPLAY_NAME", "Sartorip.")
	///// 	.UpdateProperty("_ADDRESS", "via Matteotti, 20")
	///// 	.UpdateProperty("_CAP", "20017")
	///// 	.UpdateProperty("_CITY", "Rho (MI)")
	///// 	.UpdateProperty("_CUSTOMER_DISPLAY_NAME", "Giudici S.r.l.")
	///// 	.UpdateProperty("_CUSTOMER_ADDRESS", "Via Matteotti, 123")
	///// 	.UpdateProperty("_CUSTOMER_CAP", "20017")
	///// 	.UpdateProperty("_CUSTOMER_CITY", "Rho (MI)")
	///// 	.UpdateProperty("_ADDITIONAL_INFO", null)
	///// 	.UpdateProperty("_BILLING_DATE", "01/12/2017")
	///// 	.UpdateProperty("_BILLING_NUMBER", "133")
	///// 	.AddRowsToTable<List<string>>("Tabella fattura", new List<List<string>>()
	///// 								{
	///// 											new List<string>() { "1", "My fake descr", "2.50 Euro", "2.50 Euro"},
	///// 											new List<string>() { "2", "My fake descr #2", "2.00 Euro", "4.00 Euro"},
	///// 								})
	///// 	.SaveAs("fattura01.docx");
	///// </example>
	public class DocxContentEditor
	{
		string _TargetFolder;
		string _FileName;
		private byte[] _DocumentContent;
		private Dictionary<string, object> _CustomPropertiesToAdd;
		private Dictionary<string, object> _CustomPropertiesToUpdate;
		private Dictionary<string, IEnumerable<object>> _TablesToUpdate;

		/// <summary>
		/// Default COnstructor
		/// </summary>
		public DocxContentEditor()
		{
			this._CustomPropertiesToAdd = new Dictionary<string, object>();
			this._CustomPropertiesToUpdate = new Dictionary<string, object>();
			this._TablesToUpdate = new Dictionary<string, IEnumerable<object>>();
		}
		/// <summary>
		/// Takes the input file as source
		/// </summary>
		/// <param name="path"></param>
		/// <param name="fileName"></param>
		/// <returns></returns>
		public DocxContentEditor SelectFile(string path, string fileName)
		{
			_TargetFolder = path;
			_FileName = fileName;
			return this;
		}
		/// <summary>
		/// Opens the file and reads the content
		/// </summary>
		/// <returns></returns>
		public DocxContentEditor OpenFile()
		{
			this._DocumentContent = File.ReadAllBytes(System.IO.Path.Combine(this._TargetFolder, this._FileName));
			return this;
		}
		/// <summary>
		/// Saves as file
		/// </summary>
		/// <param name="fileName"></param>
		/// <returns></returns>
		public DocxContentEditor SaveAs(string fileName)
		{
			string outputFullpath = System.IO.Path.Combine(_TargetFolder, fileName);
			if (File.Exists(outputFullpath))
				File.Delete(outputFullpath);

			using (MemoryStream stream = new MemoryStream())
			{
				// copy source file
				stream.Write(_DocumentContent, 0, (int)_DocumentContent.Length);
				//start editing
				using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
				{
					// Do work here
					var docPropertyEditor = new DocxCustomPropertiesEditor(wordDoc);
					docPropertyEditor.ReadDocProperties();

					foreach (var p in _CustomPropertiesToUpdate)
						docPropertyEditor.UpdateProperty(p.Key, p.Value as string);

					foreach (var p in _CustomPropertiesToAdd)
						docPropertyEditor.AddProperties(p.Key, p.Value as string);


					// Updating the tables
					Body bod = wordDoc.MainDocumentPart.Document.Body;

					foreach (var t in _TablesToUpdate)
					{
						// find the target
						var targetTable = bod.Descendants<Table>().SelectByDescription(t.Key);
						if (targetTable == null)
							throw new InvalidOperationException($"Unable to find table {t.Key}!!");

						foreach (var row in t.Value)
						{
							// create tableRow
							var newRow = new TableRow();

							foreach (var cell in (List<string>)row)
							{
								TableCell currentCell = new TableCell(new Paragraph(new Run(new Text(cell))));
								newRow.Append(currentCell);
							}

							// add new row
							targetTable.AppendChild(newRow);
						}
					}

					// refresh the document when it is opened
					wordDoc.UpdatePropertiesOnOpening();

				}
				// Save the file with the new name
				File.WriteAllBytes(outputFullpath, stream.ToArray());
			}
			return this;
		}
		/// <summary>
		/// Add DocProperty to file
		/// </summary>
		/// <param name="propertyName"></param>
		/// <param name="value"></param>
		/// <returns></returns>
		public DocxContentEditor AddProperty(string propertyName, string value)
		{
			if (this._CustomPropertiesToAdd.ContainsKey(propertyName))
				throw new InvalidOperationException($"Unable to add property {propertyName}: item already added");
			this._CustomPropertiesToAdd.Add(propertyName, value);
			return this;
		}
		/// <summary>
		/// Updates the value of Doc Property
		/// </summary>
		/// <param name="propertyName"></param>
		/// <param name="value"></param>
		/// <returns></returns>
		public DocxContentEditor UpdateProperty(string propertyName, string value)
		{
			if (this._CustomPropertiesToUpdate.ContainsKey(propertyName))
				throw new InvalidOperationException($"Unable to update property {propertyName}: item already added");
			this._CustomPropertiesToUpdate.Add(propertyName, value);
			return this;
		}
		/// <summary>
		/// Add rows to table
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="tableDescription"></param>
		/// <param name="rows"></param>
		/// <returns></returns>
		public DocxContentEditor AddRowsToTable<T>(string tableDescription, IEnumerable<T> rows)
		{
			if (this._TablesToUpdate.ContainsKey(tableDescription))
				throw new InvalidOperationException($"Unable to add rows to Table {tableDescription}: item already added");
			this._TablesToUpdate.Add(tableDescription, rows != null ? rows.Select(x => x as object) : null);
			return this;
		}
	}
}
