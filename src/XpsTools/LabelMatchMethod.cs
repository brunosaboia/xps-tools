namespace Std.XpsTools
{
	/// <summary>
	/// When an annotation is being positioned
	/// <c>LabelRelative</c> <c>LabelMatchMethod</c>
	/// indicates the text matching method to use.
	/// </summary>
	public enum LabelMatchMethod
	{
		/// <summary>
		/// Performs an exact case sensitive match against
		/// text in the .xps to locate the anchor label.
		/// </summary>
		ExactMatch,

		/// <summary>
		/// Performs an exact case insensitive match against
		/// text in the .xps to locate the anchor label.
		/// </summary>
		ExactMatchIgnoreCase,

		/// <summary>
		/// Indicates the anchor label text is a regular expression,
		/// which will perform a case insensitive regular expression
		/// search for the anchor label.
		/// </summary>
		RegularExpression
	}
}