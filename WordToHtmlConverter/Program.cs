using System;
using System.Xml;
using System.Xml.Linq;

using OpenXmlPowerTools;

namespace WordToHtmlConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.Error.WriteLine("Usage: WordToHtmlConverter.exe <input filename> <output filename>");
                Environment.Exit(1);
            }

            WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings();
            XElement e = WmlToHtmlConverter.ConvertToHtml(new WmlDocument(args[0]), settings);

            XmlWriterSettings s = new XmlWriterSettings();
            s.ConformanceLevel = ConformanceLevel.Fragment;
            s.OmitXmlDeclaration = true;
            s.Encoding = new System.Text.UTF8Encoding(false);
            XmlWriter writer = XmlWriter.Create(args[1], s);

            writer.WriteRaw("<style scoped>");
            writer.WriteValue(e.Element(XhtmlNoNamespace.body).Element(XhtmlNoNamespace.style).Value);
            writer.WriteRaw("</style>");
            foreach (XElement n in e.Element(XhtmlNoNamespace.body).Element(XhtmlNoNamespace.style).ElementsAfterSelf())
            {
                n.WriteTo(writer);
            }
            writer.Close();
        }
    }
}
