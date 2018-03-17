using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Umbrella.Xlsx.Helpers
{
	internal static class SpreadsheetDocumentExtension
	{
		/// <summary>
		/// Adds a new empty sheet in the document
		/// </summary>
		public static SheetData AddSheet(this SpreadsheetDocument document, string sheetName)
		{
			WorksheetPart worksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();
			var data = new SheetData();
			worksheetPart.Worksheet = new Worksheet(data);
			worksheetPart.Worksheet.Save();

			var workbookPart = document.WorkbookPart;
			// get or create colelction of sheets
			Sheets sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>() ?? workbookPart.Workbook.AppendChild(new Sheets());
			// create a new one
			Sheet sheet = new Sheet()
			{
				Id = workbookPart.GetIdOfPart(worksheetPart),
				SheetId = (UInt32)(sheets.Count()) + 1,
				Name = sheetName
			};
			// add to collection and save
			sheets.Append(sheet);
			document.WorkbookPart.Workbook.Save();

			return data;
		}
		/// <summary>
		/// Remove an existing sheet from file
		/// </summary>
		/// <param name="excelDoc"></param>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		public static bool DeleteSheet(this SpreadsheetDocument excelDoc, string sheetName)
		{
			bool returnValue = false;
			XmlDocument doc = new XmlDocument();
			doc.Load(excelDoc.WorkbookPart.GetStream());
			XmlNamespaceManager nsManager = new XmlNamespaceManager(doc.NameTable);
			nsManager.AddNamespace("d", doc.DocumentElement.NamespaceURI);
			string searchString = string.Format("//d:sheet[@name='{0}']", sheetName);
			XmlNode node = doc.SelectSingleNode(searchString, nsManager);
			if (node != null)
			{
				XmlAttribute relationAttribute = node.Attributes["r:id"];
				if (relationAttribute != null)
				{
					string relId = relationAttribute.Value;
					excelDoc.WorkbookPart.DeletePart(relId);
					node.ParentNode.RemoveChild(node);
					doc.Save(excelDoc.WorkbookPart.GetStream(FileMode.Create));
					returnValue = true;
				}
			}
			else
				throw new InvalidOperationException("Unable to find sheet " + sheetName);

			return returnValue;
		}
		/// <summary>
		/// Gets the list of sheets from a File
		/// </summary>
		/// <param name="excelDoc"></param>
		/// <returns></returns>
		public static List<string> GetSheetNames(this SpreadsheetDocument excelDoc)
		{
			List<string> sheets = new List<string>();
			WorkbookPart workbook = excelDoc.WorkbookPart;
			Stream workbookstr = workbook.GetStream();
			XmlDocument doc = new XmlDocument();
			doc.Load(workbookstr);
			XmlNamespaceManager nsManager = new XmlNamespaceManager(doc.NameTable);
			nsManager.AddNamespace("default", doc.DocumentElement.NamespaceURI);
			XmlNodeList nodelist = doc.SelectNodes("//default:sheets/default:sheet", nsManager);
			foreach (XmlNode node in nodelist)
			{
				string sheetName = string.Empty;
				sheetName = node.Attributes["name"].Value;
				sheets.Add(sheetName);
			}

			return sheets;
		}


	}
}
