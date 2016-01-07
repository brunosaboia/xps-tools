using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using Std.XpsTools;

namespace Annotator
{
	static class Program
	{
		/// <summary>
		/// 
		/// </summary>
		[STAThread]
		static void Main(string[] args)
		{
            var editor = new XpsEditor(new List<Annotation>
            {
				//create a single annotation on the first page in blue, bold and italics
				//at fixed position 100, 100
				new Annotation
                {
                    ForegroundColor = Colors.Blue,
                    PositionMethod = PositioningMethod.Absolute,
                    Position = new Point(37,240),
                    Text = "TestXpsTools",
                    TextSize = 15,
                    IsItalic = true,
                    FontWeight = FontWeights.Bold,
                    PageNumber = 0
				},

                //create annotation in each page
                new Annotation
                {
                    ForegroundColor = Colors.Brown,
                    PositionMethod = PositioningMethod.Absolute,
                    Position = new Point(150,170),
                    Text = "Writing XPS files is fun",
                    TextSize = 25
                }
            });

            try
            {
                editor.ApplyAnnotations(@"..\..\..\TestFiles\Office2007_EUROTEST.xps",
                    @"..\..\..\TestFiles\Office2007_EUROTEST-modified.xps");
            }
            catch(Exception ex)
            {
                Console.WriteLine($"Error saving {ex.Message}");
                Console.Read();
            }
		}
	}
}
