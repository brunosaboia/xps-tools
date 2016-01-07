using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Xps.Packaging;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Xps.Serialization;

namespace Std.XpsTools
{
	public class XpsEditor
	{
		private class AnnotationData
		{
			public Annotation Annotation { get; }
			public Regex LabelMatcher { get; }
			public string LabelMatchText { get; }

			public AnnotationData(Annotation annotation)
			{
				Annotation = annotation;

				if (annotation.PositionMethod == PositioningMethod.LabelRelative &&
					annotation.MatchMethod == LabelMatchMethod.RegularExpression)
				{
					LabelMatcher = new Regex(annotation.AnchorLabel,
						RegexOptions.Compiled | RegexOptions.IgnoreCase);
				}
				else if (annotation.MatchMethod == null ||
						annotation.MatchMethod == LabelMatchMethod.ExactMatch)
				{
					LabelMatchText = annotation.AnchorLabel;
				}
				else
				{
					LabelMatchText = annotation.AnchorLabel.ToLower();
				}
			}
		}

		private readonly Dictionary<int, AnnotationData[]> _pageAnnotations;
		private readonly AnnotationData[] _commonAnnotations;
		private readonly Dictionary<Color, Brush> _brushCache = new Dictionary<Color, Brush>(); 
		private static readonly AnnotationData[] _emptyAnnotations = new AnnotationData[0];
		private string _intermediateOutputPath;
		private readonly HashSet<int> _annotatedPages = new HashSet<int>();

		private const string AnnotationCanvasName = "_Annotations_";

        public XpsEditor(List<Annotation> annotations)
		{
			if (annotations == null)
			{
				throw new ArgumentNullException(nameof(annotations));
			}
			var byPages = annotations.GroupBy(a => a.PageNumber)
				.Select(g => new
				{
					Number = g.Key ?? -1,
					Annotations = g.ToArray()
				})
				.OrderBy(g => g.Number)
				.ToList();

			_commonAnnotations = _emptyAnnotations;

			var commonAnnotations = byPages.FirstOrDefault(p => p.Number == -1);
			if (commonAnnotations != null)
			{
				byPages.Remove(commonAnnotations);
				_commonAnnotations = commonAnnotations.Annotations.Select(a => new AnnotationData(a)).ToArray();
			}

			_pageAnnotations = byPages.ToDictionary(p => p.Number, p => p.Annotations
				.Select(a => new AnnotationData(a))
				.ToArray());
		}

		public int ApplyAnnotations(string sourcePath, string outputPath)
		{
			ValidatePath(sourcePath, nameof(sourcePath), mustExist: true);
			ValidatePath(outputPath, nameof(outputPath));

			_intermediateOutputPath = outputPath + ".temp";

			try
			{
				if (File.Exists(outputPath))
				{
					File.Delete(outputPath);
				}
				if (File.Exists(_intermediateOutputPath))
				{
					File.Delete(_intermediateOutputPath);
				}

				_annotatedPages.Clear();

				using (var sourceDocument = new XpsDocument(sourcePath, FileAccess.ReadWrite))
				{
					var documentSequence = sourceDocument.GetFixedDocumentSequence();
					var fixedDocument = documentSequence.References[0].GetDocument(true);

					for (var pageNumber = 0; pageNumber < fixedDocument.Pages.Count; pageNumber++)
					{
						ApplyAnnotations(fixedDocument, pageNumber);
					}

					if (_annotatedPages.Count == 0)
					{
						//no annotations were applied
						File.Copy(sourcePath, outputPath);
						return 0;
					}

					Save(_intermediateOutputPath, documentSequence);
				}

				RewriteXps(_intermediateOutputPath, outputPath);
			}
			finally
			{
				if (File.Exists(_intermediateOutputPath))
				{
					File.Delete(_intermediateOutputPath);
				}
			}

			return _annotatedPages.Count;
		}

		private class GlyphData
		{
			public Glyphs Glyphs { get; }
			public string LowerCaseText { get; }

			public GlyphData(Glyphs glyphs)
			{
				Glyphs = glyphs;
				LowerCaseText = glyphs.UnicodeString.ToLower();
			}
		}

		private readonly Transform _canvasRenderTransform = new MatrixTransform(1, 0, 0, 1, 0, 0);

        private void ApplyAnnotations(FixedDocument sourceDocument, int pageNumber)
		{
			AnnotationData[] pageAnnotations;
			_pageAnnotations.TryGetValue(pageNumber, out pageAnnotations);
			if (pageAnnotations == null &&
				_commonAnnotations == null)
			{
				//no annotations for this page
				return;
			}

			var page = sourceDocument.Pages[pageNumber].GetPageRoot(forceReload: true);

			var pageSize = new Size(page.Width, page.Height);
			page.Measure(pageSize);
			page.Arrange(new Rect(new Point(), pageSize));
			page.UpdateLayout();

			var pageLabels = page.Children.OfType<Glyphs>().Select(g => new GlyphData(g)).ToArray();

			var annotationCanvas = new Canvas
			{
				Width = page.Width,
				Height = page.Height,
				RenderTransform = _canvasRenderTransform,
                Name = AnnotationCanvasName 
			};
			page.Children.Add(annotationCanvas);

			//annotations are applied page specific first, then common
			foreach (var annotation in (pageAnnotations ?? _emptyAnnotations)
				.Concat(_commonAnnotations ?? _emptyAnnotations))
			{
				ApplyAnnotation(annotation, page, pageNumber, pageLabels, annotationCanvas);
			}
		}

		private void ApplyAnnotation(AnnotationData annot, FixedPage page, int pageNumber, GlyphData[] pageLabels, Canvas canvas)
		{
			foreach (var location in CalculateAnnotationLocations(pageLabels, annot))
			{
				_annotatedPages.Add(pageNumber);
				
				var annotationText = new TextBlock(new Run(annot.Annotation.Text)
				{
					Foreground = GetForegroundBrush(annot.Annotation.ForegroundColor),
					FontSize = annot.Annotation.TextSize,
					FontWeight = annot.Annotation.FontWeight ?? FontWeights.Normal,
					FontFamily = new FontFamily(string.IsNullOrEmpty(annot.Annotation.FontName)
						? "Arial"
						: annot.Annotation.FontName),
					FontStyle = annot.Annotation.IsItalic 
						? FontStyles.Italic
						: FontStyles.Normal
				});

				if (annot.Annotation.CustomTransform != null)
				{
					annotationText.RenderTransform = annot.Annotation.CustomTransform;
				}

				Canvas.SetLeft(annotationText, location.X);
				Canvas.SetTop(annotationText, location.Y);
				canvas.Children.Add(annotationText);
			}
		}

		private IEnumerable<Point> CalculateAnnotationLocations(GlyphData[] pageLabels, AnnotationData annot)
		{
			if (annot.Annotation.PositionMethod == PositioningMethod.Absolute)
			{
				yield return annot.Annotation.Position;
			}

			IEnumerable<GlyphData> matches;

			switch (annot.Annotation.MatchMethod)
			{
				case null:
				case LabelMatchMethod.ExactMatch:
					matches = pageLabels.Where(g => g.Glyphs.UnicodeString == annot.LabelMatchText);
					break;
				case LabelMatchMethod.ExactMatchIgnoreCase:
					matches = pageLabels.Where(g => g.LowerCaseText == annot.LabelMatchText);
					break;
				case LabelMatchMethod.RegularExpression:
					matches = pageLabels.Where(g => annot.LabelMatcher.IsMatch(g.LowerCaseText));
					break;
				default:
					throw new ArgumentOutOfRangeException("Invalid MatchMethod");
			}

			foreach (var match in matches)
			{
				var glyphs = match.Glyphs;
				var anchorPos = glyphs.RenderTransform.Transform(
					new Point(glyphs.OriginX, glyphs.OriginY));
				var pos = new Point(annot.Annotation.Position.X + anchorPos.X,
					(annot.Annotation.Position.Y + anchorPos.Y) - annot.Annotation.TextSize);

				yield return pos;
			}
		}

		private Brush GetForegroundBrush(Color? color)
		{
			if (color == null)
			{
				return Brushes.Black;
			}

			Brush result;
			if (_brushCache.TryGetValue(color.Value, out result))
			{
				return result;
			}

			result = new SolidColorBrush(color.Value);
			_brushCache.Add(color.Value, result);

			return result;
		}

		private void RewriteXps(string intermediatePath, string outputPath)
		{
			using (var intermediateDocument = new XpsDocument(intermediatePath, FileAccess.ReadWrite))
			{
				var documentSequence = intermediateDocument.GetFixedDocumentSequence();
				var sourceDocument = documentSequence.References[0].GetDocument(true);

				foreach (var pageNumber in _annotatedPages)
				{
					var page = sourceDocument.Pages[pageNumber].GetPageRoot(forceReload: true);
					var annotationContainer = page.Children.OfType<Canvas>()
						.FirstOrDefault(c => c.Name == AnnotationCanvasName);
					if (annotationContainer == null)
					{
						continue;
					}

					//the XPS writer adds an invalid xml:lang="" attriute to the <Glyphs/> elements
					//that are generated from the TextBlock's associated with annotations.
					//
					//This invalid attribute renders the .xps unreadable.

					foreach (var glyph in annotationContainer.Children
						.OfType<Glyphs>()
						.Where(glyph => glyph.Language != null && 
						string.IsNullOrEmpty(glyph.Language.IetfLanguageTag)))
					{
						glyph.Language = null;
					}
				}

				Save(outputPath, documentSequence);
			}
		}

		private void Save(string outputPath, IDocumentPaginatorSource paginator)
		{
			using (var container = Package.Open(outputPath, FileMode.Create))
			{
				using (var xpsDoc = new XpsDocument(container, CompressionOption.Maximum))
				{
					var sm = new XpsSerializationManager(new XpsPackagingPolicy(xpsDoc), false);
					sm.SaveAsXaml(paginator);
				}
			}
		}

		private void ValidatePath(string path, string argName, bool mustExist = false)
		{
			if (string.IsNullOrEmpty(path))
			{
				throw new ArgumentNullException(argName);
			}

			if (mustExist && !File.Exists(path))
			{
				throw new ArgumentException($"{path} must exist", argName);
			}
		}
	}
}