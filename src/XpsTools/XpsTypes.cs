using System.Globalization;

namespace Std.XpsTools
{
	public readonly struct Point
	{
		public double X { get; }
		public double Y { get; }

		public Point(double x, double y)
		{
			X = x;
			Y = y;
		}

		public override string ToString()
		{
			return FormattableString.Invariant($"{X},{Y}");
		}
	}

	public readonly struct Color
	{
		public byte A { get; }
		public byte R { get; }
		public byte G { get; }
		public byte B { get; }

		public Color(byte a, byte r, byte g, byte b)
		{
			A = a;
			R = r;
			G = g;
			B = b;
		}

		public static Color FromRgb(byte r, byte g, byte b)
		{
			return new Color(255, r, g, b);
		}

		public static Color FromArgb(byte a, byte r, byte g, byte b)
		{
			return new Color(a, r, g, b);
		}

		public string ToXpsHex()
		{
			return string.Format(CultureInfo.InvariantCulture, "#{0:X2}{1:X2}{2:X2}{3:X2}", A, R, G, B);
		}
	}

	public static class Colors
	{
		public static Color Black => Color.FromRgb(0, 0, 0);
		public static Color Blue => Color.FromRgb(0, 0, 255);
		public static Color Brown => Color.FromRgb(165, 42, 42);
	}

	public enum FontWeight
	{
		Normal,
		Bold
	}

	public static class FontWeights
	{
		public static FontWeight Normal => FontWeight.Normal;
		public static FontWeight Bold => FontWeight.Bold;
	}

	public readonly struct Matrix
	{
		public double M11 { get; }
		public double M12 { get; }
		public double M21 { get; }
		public double M22 { get; }
		public double OffsetX { get; }
		public double OffsetY { get; }

		public Matrix(double m11, double m12, double m21, double m22, double offsetX, double offsetY)
		{
			M11 = m11;
			M12 = m12;
			M21 = m21;
			M22 = m22;
			OffsetX = offsetX;
			OffsetY = offsetY;
		}

		public static Matrix Identity => new Matrix(1, 0, 0, 1, 0, 0);

		public Point Transform(Point point)
		{
			var x = (point.X * M11) + (point.Y * M21) + OffsetX;
			var y = (point.X * M12) + (point.Y * M22) + OffsetY;
			return new Point(x, y);
		}

		public static Matrix Multiply(Matrix left, Matrix right)
		{
			var m11 = (left.M11 * right.M11) + (left.M12 * right.M21);
			var m12 = (left.M11 * right.M12) + (left.M12 * right.M22);
			var m21 = (left.M21 * right.M11) + (left.M22 * right.M21);
			var m22 = (left.M21 * right.M12) + (left.M22 * right.M22);
			var offsetX = (left.OffsetX * right.M11) + (left.OffsetY * right.M21) + right.OffsetX;
			var offsetY = (left.OffsetX * right.M12) + (left.OffsetY * right.M22) + right.OffsetY;

			return new Matrix(m11, m12, m21, m22, offsetX, offsetY);
		}

		public string ToXpsString()
		{
			return FormattableString.Invariant($"{M11},{M12},{M21},{M22},{OffsetX},{OffsetY}");
		}
	}

	public sealed class Transform
	{
		public Matrix Matrix { get; }

		public Transform(Matrix matrix)
		{
			Matrix = matrix;
		}

		public static Transform Identity { get; } = new Transform(Matrix.Identity);

		public Point TransformPoint(Point point)
		{
			return Matrix.Transform(point);
		}
	}
}
