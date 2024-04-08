using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace DividiCapitoliWord
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Inserisci il percorso del file Word:");
            string fileWord = Console.ReadLine();

            if (!File.Exists(fileWord))
            {
                Console.WriteLine("File non trovato: " + fileWord);
                return;
            }

            using (WordprocessingDocument document = WordprocessingDocument.Open(fileWord, true))
            {
                Body body = document.MainDocumentPart.Document.Body;

                Paragraph[] paragrafi = body.Elements<Paragraph>().ToArray();

                int numeroCapitolo = 1;

                foreach (Paragraph paragrafo in paragrafi)
                {
                    if (paragrafo.ParagraphProperties != null && paragrafo.ParagraphProperties.ParagraphStyleId != null &&
                        paragrafo.ParagraphProperties.ParagraphStyleId.Val == "Heading2")
                    {
                        Console.WriteLine($"\nCapitolo {numeroCapitolo}");

                        string contenutoCapitolo = "";

                        OpenXmlElement sibling = paragrafo.NextSibling();

                        while (sibling != null && !(sibling is Paragraph && ((Paragraph)sibling).ParagraphProperties != null &&
                            ((Paragraph)sibling).ParagraphProperties.ParagraphStyleId != null &&
                            ((Paragraph)sibling).ParagraphProperties.ParagraphStyleId.Val == "Heading2"))
                        {
                            if (sibling is Paragraph)
                                contenutoCapitolo += ((Paragraph)sibling).InnerText + "\n";

                            sibling = sibling.NextSibling();
                        }

                        Console.WriteLine(contenutoCapitolo.Trim());

                        numeroCapitolo++;
                    }
                }
            }
        }
    }
}
