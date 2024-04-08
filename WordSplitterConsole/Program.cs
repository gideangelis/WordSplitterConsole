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
            // Chiedi all'utente di inserire il percorso del file Word
            Console.WriteLine("Inserisci il percorso del file Word:");
            string fileWord = Console.ReadLine();

            // Verifica se il file esiste
            if (!File.Exists(fileWord))
            {
                Console.WriteLine("File non trovato: " + fileWord);
                return;
            }

            // Apre il file Word
            using (WordprocessingDocument document = WordprocessingDocument.Open(fileWord, true))
            {
                // Recupera il corpo del documento
                Body body = document.MainDocumentPart.Document.Body;

                // Ottiene i paragrafi del documento
                Paragraph[] paragrafi = body.Elements<Paragraph>().ToArray();

                // Numero del capitolo
                int numeroCapitolo = 1;

                // Ciclo sui paragrafi
                foreach (Paragraph paragrafo in paragrafi)
                {
                    // Controlla se il paragrafo è un titolo h2
                    if (paragrafo.ParagraphProperties != null && paragrafo.ParagraphProperties.ParagraphStyleId != null &&
                        paragrafo.ParagraphProperties.ParagraphStyleId.Val == "Heading2")
                    {
                        // Stampa il titolo del capitolo
                        Console.WriteLine($"\nCapitolo {numeroCapitolo}");

                        // Resetta il contenuto del capitolo
                        string contenutoCapitolo = "";

                        OpenXmlElement sibling = paragrafo.NextSibling();

                        // Ciclo sui paragrafi successivi fino al prossimo titolo h2
                        while (sibling != null && !(sibling is Paragraph && ((Paragraph)sibling).ParagraphProperties != null &&
                            ((Paragraph)sibling).ParagraphProperties.ParagraphStyleId != null &&
                            ((Paragraph)sibling).ParagraphProperties.ParagraphStyleId.Val == "Heading2"))
                        {
                            // Aggiunge il contenuto del paragrafo al capitolo
                            if (sibling is Paragraph)
                                contenutoCapitolo += ((Paragraph)sibling).InnerText + "\n";

                            sibling = sibling.NextSibling();
                        }

                        // Stampa il contenuto del capitolo
                        Console.WriteLine(contenutoCapitolo.Trim());

                        // Incrementa il numero del capitolo
                        numeroCapitolo++;
                    }
                }
            }
        }
    }
}
