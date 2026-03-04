using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using SixLabors.Fonts;

namespace Std.XpsTools
{
	public class XpsEditor
	{
		private sealed class AnnotationData
		{
			public Annotation Annotation { get; }
			public Regex? LabelMatcher { get; }
			public string LabelMatchText { get; }

			public AnnotationData(Annotation annotation)
			{
				Annotation = annotation;
				var anchorLabel = annotation.AnchorLabel ?? string.Empty;

				if (annotation.PositionMethod == PositioningMethod.LabelRelative &&
					annotation.MatchMethod == LabelMatchMethod.RegularExpression)
				{
					LabelMatcher = new Regex(anchorLabel,
						RegexOptions.Compiled | RegexOptions.IgnoreCase);
					LabelMatchText = anchorLabel;
				}
				else if (annotation.MatchMethod == null ||
					annotation.MatchMethod == LabelMatchMethod.ExactMatch)
				{
					LabelMatchText = anchorLabel;
				}
				else
				{
					LabelMatchText = anchorLabel.ToLowerInvariant();
				}
			}
		}

		private sealed class PageInfo
		{
			public string PartName { get; }
			public int Index { get; }

			public PageInfo(string partName, int index)
			{
				PartName = partName;
				Index = index;
			}
		}

		private sealed class GlyphInfo
		{
			public string Text { get; }
			public string LowerText { get; }
			public Point Origin { get; }
			public string? FontUri { get; }
			public double? FontSize { get; }

			public GlyphInfo(string text, Point origin, string? fontUri, double? fontSize)
			{
				Text = text;
				LowerText = text.ToLowerInvariant();
				Origin = origin;
				FontUri = fontUri;
				FontSize = fontSize;
			}
		}

		private readonly Dictionary<int, AnnotationData[]> _pageAnnotations;
		private readonly AnnotationData[] _commonAnnotations;
		private readonly HashSet<int> _annotatedPages = new HashSet<int>();
		private static readonly AnnotationData[] EmptyAnnotations = Array.Empty<AnnotationData>();
		private const string AnnotationCanvasName = "_Annotations_";
		private static readonly XNamespace XpsNamespace = "http://schemas.microsoft.com/xps/2005/06";

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

			_commonAnnotations = EmptyAnnotations;
			var commonAnnotations = byPages.FirstOrDefault(p => p.Number == -1);
			if (commonAnnotations != null)
			{
				byPages.Remove(commonAnnotations);
				_commonAnnotations = commonAnnotations.Annotations.Select(a => new AnnotationData(a)).ToArray();
			}

			_pageAnnotations = byPages.ToDictionary(
				p => p.Number,
				p => p.Annotations.Select(a => new AnnotationData(a)).ToArray());
		}

		public int ApplyAnnotations(string sourcePath, string outputPath)
		{
			ValidatePath(sourcePath, nameof(sourcePath), mustExist: true);
			ValidatePath(outputPath, nameof(outputPath));

			var tempPath = outputPath + ".temp";
			if (File.Exists(outputPath))
			{
				File.Delete(outputPath);
			}
			if (File.Exists(tempPath))
			{
				File.Delete(tempPath);
			}

			_annotatedPages.Clear();

			var succeeded = false;
			try
			{
				using var sourceArchive = ZipFile.OpenRead(sourcePath);
				var pages = GetPages(sourceArchive);
				var pageIndexByPath = pages.ToDictionary(p => p.PartName, p => p.Index, StringComparer.Ordinal);
				var packageFonts = GetPackageFontUris(sourceArchive);
				var fontMap = BuildFontNameMap(sourceArchive);

				using var outputArchive = ZipFile.Open(tempPath, ZipArchiveMode.Create);
				foreach (var entry in sourceArchive.Entries)
				{
					if (pageIndexByPath.TryGetValue(entry.FullName, out var pageIndex) &&
						HasAnnotationsForPage(pageIndex))
					{
						var updated = ApplyAnnotationsToPage(entry, pageIndex, packageFonts, fontMap, out var modified);
						WriteEntry(outputArchive, entry.FullName, updated, entry.LastWriteTime);
						if (modified)
						{
							_annotatedPages.Add(pageIndex);
						}
					}
					else
					{
						CopyEntry(entry, outputArchive);
					}
				}

				succeeded = true;
			}
			finally
			{
				if (succeeded)
				{
					if (File.Exists(tempPath))
					{
						File.Move(tempPath, outputPath, overwrite: true);
					}
				}
				else if (File.Exists(tempPath))
				{
					File.Delete(tempPath);
				}
			}

			return _annotatedPages.Count;
		}

		private bool HasAnnotationsForPage(int pageIndex)
		{
			return _commonAnnotations.Length > 0 || _pageAnnotations.ContainsKey(pageIndex);
		}

		private byte[] ApplyAnnotationsToPage(
			ZipArchiveEntry entry,
			int pageIndex,
			List<string> packageFonts,
			Dictionary<string, string> fontMap,
			out bool modified)
		{
			using var sourceStream = entry.Open();
			using var buffer = new MemoryStream();
			sourceStream.CopyTo(buffer);
			var bytes = buffer.ToArray();

			var annotations = GetAnnotationsForPage(pageIndex);
			if (annotations.Length == 0)
			{
				modified = false;
				return bytes;
			}

			var encodingInfo = DetectEncoding(bytes);
			using var xmlStream = new MemoryStream(bytes);
			var doc = XDocument.Load(xmlStream, LoadOptions.PreserveWhitespace);
			var fixedPage = doc.Root;
			if (fixedPage == null)
			{
				modified = false;
				return bytes;
			}

			var glyphs = ExtractGlyphs(fixedPage, Matrix.Identity).ToArray();
			var defaultFontUri = GetDefaultFontUri(glyphs, packageFonts);
			var defaultFontSize = glyphs.Select(g => g.FontSize).FirstOrDefault(size => size.HasValue) ?? 12d;

			var addedGlyphs = new List<XElement>();
			foreach (var annotation in annotations)
			{
				var fontSize = annotation.Annotation.TextSize > 0
					? annotation.Annotation.TextSize
					: defaultFontSize;

				foreach (var location in CalculateAnnotationLocations(glyphs, annotation, fontSize))
				{
					addedGlyphs.Add(CreateGlyphElement(annotation, location, defaultFontUri, fontSize, fontMap));
				}
			}

			if (addedGlyphs.Count == 0)
			{
				modified = false;
				return bytes;
			}

			var canvas = new XElement(XpsNamespace + "Canvas",
				new XAttribute("Name", AnnotationCanvasName));
			foreach (var glyph in addedGlyphs)
			{
				canvas.Add(glyph);
			}
			fixedPage.Add(canvas);

			using var outputStream = new MemoryStream();
			var writerSettings = new XmlWriterSettings
			{
				Encoding = encodingInfo.Encoding,
				Indent = false,
				OmitXmlDeclaration = doc.Declaration == null
			};
			using (var writer = XmlWriter.Create(outputStream, writerSettings))
			{
				doc.Save(writer);
			}
			modified = true;
			return outputStream.ToArray();
		}

		private AnnotationData[] GetAnnotationsForPage(int pageNumber)
		{
			_pageAnnotations.TryGetValue(pageNumber, out var pageAnnotations);
			if (pageAnnotations == null && _commonAnnotations.Length == 0)
			{
				return EmptyAnnotations;
			}

			return (pageAnnotations ?? EmptyAnnotations)
				.Concat(_commonAnnotations)
				.ToArray();
		}

		private static IEnumerable<GlyphInfo> ExtractGlyphs(XElement element, Matrix inheritedTransform)
		{
			var currentTransform = Matrix.Multiply(inheritedTransform, GetRenderTransform(element));

			if (element.Name == XpsNamespace + "Glyphs")
			{
				var text = element.Attribute("UnicodeString")?.Value ?? string.Empty;
				var originX = ParseDouble(element.Attribute("OriginX")?.Value);
				var originY = ParseDouble(element.Attribute("OriginY")?.Value);
				var origin = currentTransform.Transform(new Point(originX, originY));
				var fontUri = element.Attribute("FontUri")?.Value;
				var fontSize = ParseOptionalDouble(element.Attribute("FontRenderingEmSize")?.Value);
				yield return new GlyphInfo(text, origin, fontUri, fontSize);
			}

			foreach (var child in element.Elements())
			{
				if (child.Name.LocalName.EndsWith(".RenderTransform", StringComparison.Ordinal))
				{
					continue;
				}
				foreach (var glyph in ExtractGlyphs(child, currentTransform))
				{
					yield return glyph;
				}
			}
		}

		private static Matrix GetRenderTransform(XElement element)
		{
			var attr = element.Attribute("RenderTransform");
			if (attr != null)
			{
				return ParseMatrix(attr.Value);
			}

			var renderTransform = element.Elements()
				.FirstOrDefault(e => e.Name.LocalName.EndsWith(".RenderTransform", StringComparison.Ordinal));
			if (renderTransform == null)
			{
				return Matrix.Identity;
			}

			var matrixTransform = renderTransform.Element(XpsNamespace + "MatrixTransform");
			var matrixValue = matrixTransform?.Attribute("Matrix")?.Value;
			return ParseMatrix(matrixValue);
		}

		private static Matrix ParseMatrix(string? value)
		{
			if (string.IsNullOrWhiteSpace(value))
			{
				return Matrix.Identity;
			}

			var trimmed = value.Trim();
			if (trimmed.StartsWith("matrix", StringComparison.OrdinalIgnoreCase))
			{
				var start = trimmed.IndexOf('(');
				var end = trimmed.LastIndexOf(')');
				if (start >= 0 && end > start)
				{
					trimmed = trimmed.Substring(start + 1, end - start - 1);
				}
			}

			var parts = trimmed.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
			if (parts.Length != 6)
			{
				return Matrix.Identity;
			}

			return new Matrix(
				ParseDouble(parts[0]),
				ParseDouble(parts[1]),
				ParseDouble(parts[2]),
				ParseDouble(parts[3]),
				ParseDouble(parts[4]),
				ParseDouble(parts[5]));
		}

		private static double ParseDouble(string? value)
		{
			if (string.IsNullOrWhiteSpace(value))
			{
				return 0d;
			}

			if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result))
			{
				return result;
			}

			return 0d;
		}

		private static double? ParseOptionalDouble(string? value)
		{
			if (string.IsNullOrWhiteSpace(value))
			{
				return null;
			}

			if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result))
			{
				return result;
			}

			return null;
		}

		private IEnumerable<Point> CalculateAnnotationLocations(GlyphInfo[] pageLabels, AnnotationData annotation, double fontSize)
		{
			if (annotation.Annotation.PositionMethod == PositioningMethod.Absolute)
			{
				yield return annotation.Annotation.Position;
			}

			IEnumerable<GlyphInfo> matches;
			switch (annotation.Annotation.MatchMethod)
			{
				case null:
				case LabelMatchMethod.ExactMatch:
					matches = pageLabels.Where(g => g.Text == annotation.LabelMatchText);
					break;
				case LabelMatchMethod.ExactMatchIgnoreCase:
					matches = pageLabels.Where(g => g.LowerText == annotation.LabelMatchText);
					break;
				case LabelMatchMethod.RegularExpression:
					matches = pageLabels.Where(g => annotation.LabelMatcher != null &&
						annotation.LabelMatcher.IsMatch(g.LowerText));
					break;
				default:
					throw new ArgumentOutOfRangeException(nameof(annotation.Annotation.MatchMethod),
						"Invalid MatchMethod");
			}

			foreach (var match in matches)
			{
				var anchorPos = match.Origin;
				var pos = new Point(
					annotation.Annotation.Position.X + anchorPos.X,
					(annotation.Annotation.Position.Y + anchorPos.Y) - fontSize);
				yield return pos;
			}
		}

		private XElement CreateGlyphElement(
			AnnotationData annotation,
			Point position,
			string defaultFontUri,
			double fontSize,
			Dictionary<string, string> fontMap)
		{
			var fontUri = ResolveFontUri(annotation.Annotation.FontName, defaultFontUri, fontMap);
			var glyph = new XElement(XpsNamespace + "Glyphs",
				new XAttribute("UnicodeString", annotation.Annotation.Text ?? string.Empty),
				new XAttribute("OriginX", FormatDouble(position.X)),
				new XAttribute("OriginY", FormatDouble(position.Y)),
				new XAttribute("FontRenderingEmSize", FormatDouble(fontSize)),
				new XAttribute("FontUri", fontUri));

			var color = annotation.Annotation.ForegroundColor ?? Colors.Black;
			glyph.SetAttributeValue("Fill", color.ToXpsHex());

			var style = GetStyleSimulations(annotation.Annotation.FontWeight, annotation.Annotation.IsItalic);
			if (!string.IsNullOrEmpty(style))
			{
				glyph.SetAttributeValue("StyleSimulations", style);
			}

			if (annotation.Annotation.CustomTransform != null)
			{
				var renderTransform = new XElement(XpsNamespace + "Glyphs.RenderTransform",
					new XElement(XpsNamespace + "MatrixTransform",
						new XAttribute("Matrix", annotation.Annotation.CustomTransform.Matrix.ToXpsString())));
				glyph.Add(renderTransform);
			}

			return glyph;
		}

		private static string ResolveFontUri(string? fontName, string defaultFontUri, Dictionary<string, string> fontMap)
		{
			if (!string.IsNullOrWhiteSpace(fontName) && fontMap.TryGetValue(fontName, out var mapped))
			{
				return mapped;
			}

			return defaultFontUri;
		}

		private static string? GetStyleSimulations(FontWeight? fontWeight, bool isItalic)
		{
			var bold = fontWeight == FontWeight.Bold;
			if (bold && isItalic)
			{
				return "BoldItalicSimulation";
			}
			if (bold)
			{
				return "BoldSimulation";
			}
			if (isItalic)
			{
				return "ItalicSimulation";
			}
			return null;
		}

		private static string GetDefaultFontUri(GlyphInfo[] glyphs, List<string> packageFonts)
		{
			var fromPage = glyphs.Select(g => g.FontUri).FirstOrDefault(uri => !string.IsNullOrWhiteSpace(uri));
			if (!string.IsNullOrWhiteSpace(fromPage))
			{
				return fromPage!;
			}

			var fromPackage = packageFonts.FirstOrDefault();
			if (!string.IsNullOrWhiteSpace(fromPackage))
			{
				return fromPackage;
			}

			throw new InvalidOperationException("No font resources were found in the XPS package.");
		}

		private static List<string> GetPackageFontUris(ZipArchive archive)
		{
			return archive.Entries
				.Where(e => e.FullName.StartsWith("Resources/Fonts/", StringComparison.OrdinalIgnoreCase))
				.Select(e => "/" + e.FullName)
				.ToList();
		}

		private static Dictionary<string, string> BuildFontNameMap(ZipArchive archive)
		{
			var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
			var collection = new FontCollection();
			foreach (var entry in archive.Entries.Where(e =>
				e.FullName.StartsWith("Resources/Fonts/", StringComparison.OrdinalIgnoreCase)))
			{
				try
				{
					using var stream = entry.Open();
					using var buffer = new MemoryStream();
					stream.CopyTo(buffer);
					buffer.Position = 0;

					var family = collection.Add(buffer);
					if (!map.ContainsKey(family.Name))
					{
						map[family.Name] = "/" + entry.FullName;
					}
				}
				catch
				{
					//best-effort font inspection; ignore unparseable or obfuscated fonts
				}
			}

			return map;
		}

		private static List<PageInfo> GetPages(ZipArchive archive)
		{
			var fixedDocSeqEntry = archive.GetEntry("FixedDocSeq.fdseq");
			if (fixedDocSeqEntry == null)
			{
				throw new InvalidOperationException("FixedDocSeq.fdseq was not found in the XPS package.");
			}

			using var seqStream = fixedDocSeqEntry.Open();
			var seqDoc = XDocument.Load(seqStream, LoadOptions.PreserveWhitespace);
			var documentRefs = seqDoc.Root?.Elements(XpsNamespace + "DocumentReference") ?? Enumerable.Empty<XElement>();

			var pages = new List<PageInfo>();
			var index = 0;
			foreach (var docRef in documentRefs)
			{
				var docPath = docRef.Attribute("Source")?.Value;
				if (string.IsNullOrWhiteSpace(docPath))
				{
					continue;
				}

				var docEntry = archive.GetEntry(docPath);
				if (docEntry == null)
				{
					continue;
				}

				using var docStream = docEntry.Open();
				var doc = XDocument.Load(docStream, LoadOptions.PreserveWhitespace);
				var pageContents = doc.Root?.Elements(XpsNamespace + "PageContent") ?? Enumerable.Empty<XElement>();
				var basePath = GetDirectoryName(docPath);

				foreach (var pageContent in pageContents)
				{
					var pageSource = pageContent.Attribute("Source")?.Value;
					if (string.IsNullOrWhiteSpace(pageSource))
					{
						continue;
					}

					var pagePath = CombineXpsPath(basePath, pageSource);
					pages.Add(new PageInfo(pagePath, index));
					index++;
				}
			}

			return pages;
		}

		private static string GetDirectoryName(string path)
		{
			var lastSlash = path.LastIndexOf('/');
			if (lastSlash <= 0)
			{
				return string.Empty;
			}
			return path.Substring(0, lastSlash);
		}

		private static string CombineXpsPath(string basePath, string relativePath)
		{
			var trimmedBase = basePath?.TrimEnd('/') ?? string.Empty;
			var trimmedRelative = relativePath?.TrimStart('/') ?? string.Empty;
			if (string.IsNullOrEmpty(trimmedBase))
			{
				return trimmedRelative;
			}
			return $"{trimmedBase}/{trimmedRelative}";
		}

		private static string FormatDouble(double value)
		{
			return value.ToString("G", CultureInfo.InvariantCulture);
		}

		private static void CopyEntry(ZipArchiveEntry sourceEntry, ZipArchive outputArchive)
		{
			using var sourceStream = sourceEntry.Open();
			var entry = outputArchive.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
			entry.LastWriteTime = sourceEntry.LastWriteTime;
			using var destStream = entry.Open();
			sourceStream.CopyTo(destStream);
		}

		private static void WriteEntry(ZipArchive outputArchive, string entryName, byte[] content, DateTimeOffset timestamp)
		{
			var entry = outputArchive.CreateEntry(entryName, CompressionLevel.Optimal);
			entry.LastWriteTime = timestamp;
			using var destStream = entry.Open();
			destStream.Write(content, 0, content.Length);
		}

		private static EncodingInfo DetectEncoding(byte[] bytes)
		{
			if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
			{
				return new EncodingInfo(new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
			}
			if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
			{
				return new EncodingInfo(new UnicodeEncoding(bigEndian: false, byteOrderMark: true));
			}
			if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
			{
				return new EncodingInfo(new UnicodeEncoding(bigEndian: true, byteOrderMark: true));
			}

			return new EncodingInfo(new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
		}

		private sealed class EncodingInfo
		{
			public Encoding Encoding { get; }

			public EncodingInfo(Encoding encoding)
			{
				Encoding = encoding;
			}
		}

		private void ValidatePath(string path, string argName, bool mustExist = false)
		{
			if (string.IsNullOrWhiteSpace(path))
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
