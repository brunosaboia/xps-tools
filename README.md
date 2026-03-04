# xps-tools
A tool written in C# for writing annotations on XPS documents

This tool tries to make it easy to write annotations to an existing [XPS file](https://en.wikipedia.org/wiki/Open_XML_Paper_Specification). 

## .NET 10
This version targets .NET 10 and runs cross-platform by updating XPS page markup directly.
Fonts for annotations are selected from fonts already embedded in the XPS package.
When `FontName` is provided, the editor uses SixLabors.Fonts to match it to an
embedded font and falls back to the page's default font when no match is found.

# License
This project is licensed  under the [MIT License](http://opensource.org/licenses/mit-license.php).
