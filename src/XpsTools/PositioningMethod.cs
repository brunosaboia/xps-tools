namespace Std.XpsTools
{
	/// <summary>
	/// Identifies the method of positioning the annotation.
	/// </summary>
	public enum PositioningMethod
	{
		/// <summary>
		/// Invalid value - a property with this value
		/// indicates it was never configured and is not valid.
		/// </summary>
		Unknown,

		/// <summary>
		/// Indicates the annotation should be positioned
		/// relative to an existing label. The annotation
		/// <c>Position</c> member indicates the offset of
		/// the annotation from the upper left corner of the label.
		/// </summary>
		LabelRelative,

		/// <summary>
		/// Indicates the annotation should be positioned
		/// relative to the upper left corner of the page.
		/// The annotation <c>Position</c> member indicates the 
		/// absolute position on the page.
		/// </summary>
		Absolute
	}
}