using System;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;

namespace Umbrella.Docx
{

	internal class DocxCustomPropertiesEditor
	{

		WordprocessingDocument _Document;
		CustomFilePropertiesPart _CustomProperties;

		public WordprocessingDocument Document { get { return _Document; } }

		public DocxCustomPropertiesEditor(WordprocessingDocument docx)
		{
			_Document = docx;
		}

		public DocxCustomPropertiesEditor ReadDocProperties()
		{
			if (_Document.CustomFilePropertiesPart == null)
			{
				_CustomProperties = _Document.AddCustomFilePropertiesPart();
				_CustomProperties.Properties = new Properties();
			}
			else
			{
				_CustomProperties = _Document.CustomFilePropertiesPart;
			}
			return this;
		}
		/// <summary>
		/// Set custom property for proper control
		/// </summary>
		/// <param name="propertyName">Custom property identificator</param>
		/// <param name="model">Data model that is needed to be stored</param>
		public DocxCustomPropertiesEditor AddProperties(string propertyName, string model)
		{
			if (_CustomProperties == null)
				throw new InvalidOperationException($"Unable to add Property:please read document before do this action");

			var newProperty = new CustomDocumentProperty();
			// every property should have the same formatId
			newProperty.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
			newProperty.Name = propertyName;
			newProperty.VTLPWSTR = new VTLPWSTR(model);
			var properties = _CustomProperties.Properties;
			if (properties != null)
			{
				// msdn sayes that propertyId should be started from 2
				int propertyId = properties.Count() == 0 ? 2 : properties.Select(item => ((CustomDocumentProperty)item).PropertyId.Value).Max() + 1;
				newProperty.PropertyId = propertyId;
				properties.AppendChild(newProperty);
				properties.Save();
			}
			return this;
		}
		/// <summary>
		/// Set custom property for proper control
		/// </summary>
		/// <param name="propertyName">Custom property identificator</param>
		/// <param name="model">Data model that is needed to be stored</param>
		public DocxCustomPropertiesEditor UpdateProperty(string propertyName, string model)
		{
			if (_CustomProperties == null)
				throw new InvalidOperationException($"Unable to update Property: please read document before do this action");

			var property = (CustomDocumentProperty)_CustomProperties.Properties.FirstOrDefault(item => ((CustomDocumentProperty)item).Name.Value == propertyName);
			if (property == null)
				throw new ArgumentException($"Doc property {propertyName} not found");
			property.VTLPWSTR = new VTLPWSTR(model);
			_CustomProperties.Properties.Save();
			return this;
		}
		/// <summary>
		/// Parse property value
		/// </summary>
		/// <param name="propertyName">Name of property which value needed to parse</param>
		/// <returns>data model with parsed data</returns>
		public string GetCustomProperties(string propertyName)
		{
			if (_CustomProperties == null)
				throw new InvalidOperationException($"Unable to get Property: please read document before do this action");


			var property = _CustomProperties.Properties.FirstOrDefault(item => ((CustomDocumentProperty)item).Name.Value == propertyName);
			if (property != null)
				return property.InnerText;//return JsonConvert.DeserializeObject<T>(property.InnerText);
			return null;
		}
		/// <summary>
		/// Remove custom property of proper element
		/// </summary>
		/// <param name="propertyName">name of the property</param>
		public DocxCustomPropertiesEditor RemoveProperty(string propertyName)
		{
			var props = _CustomProperties.Properties;
			if (props != null)
			{
				var prop = (CustomDocumentProperty)props?.FirstOrDefault(p => ((CustomDocumentProperty)p).Name.Value == propertyName);
				prop?.Remove();
			}
			return this;
		}

		public void Save(string fileName)
		{

		}
	}
}
