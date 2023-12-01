using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

//Open the file as Stream
FileStream docStream = new FileStream(@"input-doc-used.docx", FileMode.Open, FileAccess.Read);
////Loads file stream into Word document
WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx);
wordDocument.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
//Instantiation of DocIORenderer for Word to PDF conversion
DocIORenderer render = new DocIORenderer();
//render.Settings.ChartRenderingOptions.ImageFormat = Syncfusion.OfficeChart.ExportImageFormat.Png;
//Converts Word document into PDF document
PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);
//Saves the PDF file
FileStream outputPDFStream = new FileStream(@"Output.pdf", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
pdfDocument.Save(outputPDFStream);
//Close and release the streams
pdfDocument.Close();
render.Dispose();
wordDocument.Dispose();
outputPDFStream.Dispose();

static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
{

}