using System;
using System.Linq;
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

            XElement style = e.Element(XhtmlNoNamespace.body).Element(XhtmlNoNamespace.style);
            XElement article = e.Element(XhtmlNoNamespace.body).Element(XhtmlNoNamespace.div);
            XElement footnotes = article.ElementsAfterSelf().Last();

            // The scoped attribute is a "boolean attribute," so it's not supposed to have a value.
            // We have to write it manually to accomplish that.
            writer.WriteRaw("<style scoped>");
            writer.WriteValue(style.Value);
            writer.WriteRaw("</style>");

            bool wroteBreak = false;
            writer.WriteStartElement("div");           
            foreach (XElement n in article.Elements())
            {
                if (!wroteBreak)
                {
                    if (n.Name == XhtmlNoNamespace.p && (n.Attribute("class").Value == "pt-SubHead1" || n.Attribute("class").Value == "pt-Document"))
                    {
                        writer.WriteStartElement("p");
                        writer.WriteAttributeString("style", "text-align:center;");
                        writer.WriteStartElement("a");
                        writer.WriteAttributeString("href", "#begin");
                        writer.WriteRaw("&#9660; Continue Reading &#9660;");
                        writer.WriteFullEndElement();
                        writer.WriteFullEndElement();
                        writer.WriteStartElement("div");
                        writer.WriteAttributeString("style", "height:50vh");
                        writer.WriteFullEndElement();
                        writer.WriteStartElement("p");
                        writer.WriteAttributeString("id", "begin");
                        writer.WriteFullEndElement();

                        wroteBreak = true;
                    }
                }
                n.WriteTo(writer);
            }
            writer.WriteFullEndElement();

            writer.WriteRaw("<hr>");
            footnotes.WriteTo(writer);
            writer.Close();
        }
    }
}
