using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;

namespace ExtractDocumentProperties
{
    class Program
    {
        public static Dictionary<string, string> GetCustomPropertiesOfWordDocument(string filename)
        {
            using (var package = WordprocessingDocument.Open(filename, false))
            {
                var properties = new Dictionary<string, string>();
                foreach (var property in package.CustomFilePropertiesPart.Properties.Elements<CustomDocumentProperty>())
                {
                    var value = property.VTLPWSTR == null ? "" : property.VTLPWSTR.Text;
                    properties.Add(property.Name, value);
                }
                return properties;
            }
        }

        static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine("Invalid arguments! Usage: ExtractDocumentProperties filename.docx");
                return;
            }

            var filename = args[0];
            foreach (var prop in GetCustomPropertiesOfWordDocument(filename))
            {
                Console.WriteLine("{0}:{1}:{2}", filename, prop.Key, prop.Value);
            }
        }
    }
}
