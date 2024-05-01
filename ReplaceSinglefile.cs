using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;

namespace BlazorApp1.Services
{
    public class WordTemplateService
    {
        public void InsertTextAtBookmark(string templatePath, string outputPath, string content, string bookmarkName)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(templatePath, true))
            {
                // Find the bookmark by name
                var bookmarks = wordDoc.MainDocumentPart.RootElement.Descendants<BookmarkStart>()
                    .Where(b => b.Name == bookmarkName);

                var bookmark = bookmarks.FirstOrDefault();
                if (bookmark != null)
                {
                    // Find the parent paragraph of the bookmark
                    var parentParagraph = bookmark.Parent;

                    // Clear all runs in the parent paragraph
                    parentParagraph.RemoveAllChildren<Run>();

                    // Create a new run with the specified font and size
                    Run newRun = new Run();
                    newRun.AppendChild(new RunProperties(new RunFonts() { Ascii = "Tahoma" }));
                    newRun.AppendChild(new FontSize() { Val = "24" }); // 12 point size
                    newRun.AppendChild(new Text(content));

                    // Add the new run to the parent paragraph
                    parentParagraph.AppendChild(newRun);

                    // Save the changes to the output file
                    wordDoc.MainDocumentPart.Document.Save();

                    // Close the document
                    wordDoc.Close();

                    // If you want to save as a new file, you need to handle it like this:
                    File.Copy(templatePath, outputPath, overwrite: true);
                }
            }
        }
    }
}
