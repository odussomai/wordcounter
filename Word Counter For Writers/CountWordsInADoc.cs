using DocumentFormat.OpenXml.Packaging;

namespace Word_Counter_For_Writers
{
    internal class CountWordsInADoc : ICountWordsInADoc
    {

        public int Execute(string fileName)
        {
            int words;

            using (var document = WordprocessingDocument.Open(fileName, false))
            {
                words = int.Parse(document.ExtendedFilePropertiesPart.Properties.Words.Text);
            }

            return words;
        }
    }
}