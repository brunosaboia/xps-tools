using System.Windows;
using System.Windows.Media;

namespace Std.XpsTools
{
	public class Annotation
	{
		/// <summary>
		/// The text to be applied.
		/// </summary>
		public string Text { get; set; }

		/// <summary>
		/// Indicates the page number this annotation should be applied to.
		/// <c>PageNumber</c> is optional, and if not specified
		/// the annotation will be applied to every page in the document.
		/// </summary>
		public int? PageNumber { get; set; }

		/// <summary>
		/// The positioning method to use for 
		/// determining the position of this annotation.
		/// </summary>
		public PositioningMethod PositionMethod { get; set; }

		/// <summary>
		/// If <c>LabelRelative</c> positioning is used
		/// <c>AnchorLabel</c> identifies the text to
		/// use as the basis for anchoring this annotation.
		/// </summary>
		public string AnchorLabel { get; set; }

		/// <summary>
		/// If <c>LabelRelative</c> positioning is used
		/// <c>MatchMethod</c> identifies the method to use
		/// to locate the label text inside a page.
		/// 
		/// <c>MatchMethod</c> is optional, and if not specified
		/// defaults to <c>ExactMatch</c>.
		/// </summary>
		public LabelMatchMethod? MatchMethod { get; set; }

		/// <summary>
		/// Position of the annotation.
		/// For <c>LabelRelative</c> anchors the position
		/// is relative to the starting position of the label.
		/// For <c>Absolute</c> anchors the position
		/// is relative to the upper left corner of the page.
		/// </summary>
		public Point Position { get; set; }

		/// <summary>
		/// An optional transformation that if provided
		/// will be applied to the annotation.
		/// </summary>
		public Transform CustomTransform { get; set; }

		/// <summary>
		/// An optional foreground color to be aplied
		/// to the annotation text. If not specified the
		/// default is Black.
		/// </summary>
		public Color? ForegroundColor { get; set; }

		/// <summary>
		/// Size of the annotation text.
		/// 
		/// <c>TextSize</c> is specified in units of
		/// 1/96 per inch.
		/// </summary>
		public double TextSize { get; set; }

		/// <summary>
		/// The font weight of the font used to render
		/// the annotation text.
		/// 
		/// <c>TextWeight</c> is optional. If not provided
		/// the default is <c>FontWeight.Normal</c>.
		/// </summary>
		public FontWeight? FontWeight { get; set; }

		/// <summary>
		/// Name of the font used to render
		/// the annotation texxt.
		/// 
		/// <c>FontName</c> is optional. If not provided
		/// the default is Arial.
		/// </summary>
		public string FontName { get; set; }

		/// <summary>
		/// Indicates that the annotation text should
		/// be rendered italicized.
		/// </summary>
		public bool IsItalic { get; set; }
	}
}