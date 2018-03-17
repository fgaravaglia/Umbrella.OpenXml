using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Umbrella.Docx.Helpers
{
	internal static class WordProcessingExtensions
	{
		public static WordprocessingDocument CreateAnEmptyDocx(string filepath)
		{
			// Create a document by supplying the filepath. 
			using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
			{
				// Add a main document part. 
				MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

				// Create the document structure and add some text.
				mainPart.Document = new Document();
				Body body = mainPart.Document.AppendChild(new Body());
				//Paragraph para = body.AppendChild(new Paragraph());
				//Run run = para.AppendChild(new Run());
				//run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
				return wordDocument;
			}
		}

		public static WordprocessingDocument UpdatePropertiesOnOpening(this WordprocessingDocument wordDoc)
		{
			DocumentSettingsPart settingsPart = wordDoc.MainDocumentPart.GetPartsOfType<DocumentSettingsPart>().First();

			//Update Fields
			UpdateFieldsOnOpen updateFields = new UpdateFieldsOnOpen();
			updateFields.Val = new OnOffValue(true);

			settingsPart.Settings.PrependChild<UpdateFieldsOnOpen>(updateFields);
			settingsPart.Settings.Save();

			return wordDoc;
		}


		public static Table SelectByDescription(this IEnumerable<Table> tables, string descr)
		{
			foreach (var t in tables.Where(x => x.FirstChild as TableProperties != null && x.ChildElements != null))
			{
				var properties = t.FirstChild as TableProperties;
				var tableDescription = properties.ChildElements.Where(el => el is TableDescription).Select(x => (TableDescription)x).FirstOrDefault();
				if (tableDescription != null && tableDescription.Val == descr)
					return t;
			}
			return null;
		}
	}
}
