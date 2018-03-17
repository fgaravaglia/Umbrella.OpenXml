namespace Umbrella.Xlsx.DTOs
{
	/// <summary>
	/// Desctiptor for style
	/// </summary>
	public class Style
	{
		/// <summary>
		/// Size
		/// </summary>
		public double FontSize { get; set; }
		/// <summary>
		/// TRUE means the text is in bold
		/// </summary>
		public bool Bold { get; set; }
		/// <summary>
		/// text Color
		/// </summary>
		public string RbgColor { get; set; }
	}
}
