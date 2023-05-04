using Laserfiche.DocumentServices;
using Laserfiche.RepositoryAccess;
using PdfSharpCore;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using System.Drawing;
using System.Drawing.Imaging;
using Document = Laserfiche.RepositoryAccess.Document;


namespace Laserfiche_Download_Issues
{
    // internal sealed class
    public class DocumentDownload : IDocumentDownload
    {
        private string _laserficheUsername;
        private string _laserfichePassword;
        private string _laserficheServerName;
        private string _laserficheRepoName;
        private string _laserficheViewUrl;
        private string _laserficheWebLinkUrl;
        private string _laserfichePublicWebLinkUrl;
        private Session laserficheSession;

        // TODO: Replace with your Laserfiche document that is NOT an electronic document.
        //       This code is for downloading a document comprised of images that have PUBLIC digital annoation stamps applied to them.
        private int entryId = 189271;

        private string currentDirectory = Path.Combine(Directory.GetCurrentDirectory(), "..\\..\\..");

        public DocumentDownload()
        {
            this._laserficheServerName = "Your Laserfiche Config";
            this._laserficheRepoName = "Your Laserfiche Config";
            this._laserficheViewUrl = "Your Laserfiche Config";
            this._laserficheWebLinkUrl = "Your Laserfiche Config";
            this._laserfichePublicWebLinkUrl = "Your Laserfiche Config";

            this._laserficheUsername = "Your Laserfiche Creds";
            this._laserfichePassword = "Your Laserfiche Creds";

            this.currentDirectory = Path.Combine(Directory.GetCurrentDirectory(), "..\\..\\..");

            Console.WriteLine($"LaserficheService running with username: {this._laserficheUsername}");
        }

        public IEnumerable<string> DownloadLaserficheDocument()
        {
            List<string> fileNames = new List<string>();

            try
            {
                // Reuse the same session for all document download tests
                using (Session session = new Session())
                {
                    if (null == this.laserficheSession
                        || !this.laserficheSession.IsConnected
                        || !this.laserficheSession.IsAuthenticated)
                    {
                        try
                        {
                            RepositoryRegistration repository = new RepositoryRegistration(this._laserficheServerName, this._laserficheRepoName);
                            this.laserficheSession = new Session();
                            this.laserficheSession.LogIn(this._laserficheUsername, this._laserfichePassword, repository);

                            Console.WriteLine();
                            Console.WriteLine($"Downloading Laserfiche Document - EntryId: {this.entryId}");
                            Console.WriteLine("...");
                            // Now call all of the different APIs for downloading Laserfiche documents as PDF files
                            fileNames.Add(ExportViaCustomCode(this.laserficheSession, "CustomSample", this.entryId));
                            fileNames.Add(ExportViaDocumentExporter(this.laserficheSession, "DocumentExportSample", this.entryId));
                            Console.WriteLine();

                            // I thought someone mentioned using the PdfServices.dll library for being able to download pdf images
                            // however, I don't see this functionality in the library.
                            fileNames.Add(ExportViaPdfExporter(this.laserficheSession, "PdfExportSample", this.entryId));
                        }
                        finally
                        {
                            if (session != null)
                            {
                                session.Close();
                            }
                        }
                    }

                    if (session != null)
                    {
                        try
                        {
                            session.Close();
                        }
                        catch
                        {
                            // Can't do anything, swallow the exception
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
            }
            
            return fileNames;
        }

        /// <summary>
        /// Example code for how we are downloading PDF files from a Laserfiche repository that are NOT electronic documents.
        /// These PDF files are composed of images. There are PUBLIC digital annotation stamps on some of these PDF pages.
        /// We need to understand how to download PDFs from Laserfiche with the digital annotation stamps included.
        /// </summary>
        /// <param name="session"></param>
        /// <param name="filename"></param>
        /// <param name="entryId"></param>
        /// <returns></returns>
        private static string ExportViaCustomCode(Session session, string filename, int entryId)
        {
            string fullFilename = $"{filename}.pdf";
            try
            {
                string currentDirectory = Path.Combine(Directory.GetCurrentDirectory(), "..\\..\\..");
                currentDirectory = Path.GetFullPath(currentDirectory);

                string pdfPath = $"{currentDirectory}\\{fullFilename}";
                pdfPath = Path.GetFullPath(pdfPath);

                File.Delete(pdfPath);  // Clean up files from previous execution

                Console.WriteLine($"Using download pdfPath: {pdfPath}...");

                using (PdfDocument pdf = new PdfDocument())
                {
                    using (DocumentInfo docInfo = Document.GetDocumentInfo(entryId, session))
                    {
                        if (docInfo.IsElectronicDocument)
                        {
                            // This is only here for example code.
                            // If your entryId is for an Image PDF file with digital annotations,
                            // This line of code should never be executed.
                            DownloadElectronicDocument(currentDirectory, docInfo);
                        }
                        else
                        {
                            // This is the code we need help with.
                            // It does not include the PUBLIC digital annotation stamps.
                            // EntryId: 189271 is an image pdf file with stamps!
                            DownloadImagePdfDocument(currentDirectory, pdf, docInfo);
                        }

                        pdf.Save(pdfPath);
                        pdf.Close();

                        if (!File.Exists(pdfPath))
                        {
                            Console.WriteLine($"Failed to write PDF document: {docInfo.Id} to: {pdfPath}");
                        }
                    }
                }

                Console.WriteLine($"Finished exporting PDF via custom code: {fullFilename}");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Failed exporting PDF via custom code: {fullFilename}");
                Console.WriteLine(ex.Message);
            }
            return fullFilename;
        }

        // This Document Exporter works great running in a script on the Laserfiche Dev server.
        // This successfully downloads the PDF files with all of the digital annotation stamps included on the images.
        // However, we cannot run this example on our application servers because of the dependency on .net 4.0.0.
        private static string ExportViaDocumentExporter(Session session, string filename, int entryId)
        {
            string fullFilename = $"{filename}.pdf";
            try
            {
                string currentDirectory = Path.Combine(Directory.GetCurrentDirectory(), "..\\..\\..");
                currentDirectory = Path.GetFullPath(currentDirectory);

                string pdfPath = $"{currentDirectory}\\{fullFilename}";
                pdfPath = Path.GetFullPath(pdfPath);

                File.Delete(pdfPath);  // Clean up files from previous execution

                Console.WriteLine($"Using download pdfPath: {pdfPath}...");
                using (DocumentInfo docInfo = Document.GetDocumentInfo(entryId, session))
                {
                    // initialize an instance of DocumentExporter
                    DocumentExporter exporter = new DocumentExporter();
                    // configure the pdfExporter to export images as JPEG, include annotations, and burn-in redactions.
                    exporter.IncludeAnnotations = true;
                    exporter.BlackoutRedactions = true;
                    exporter.PageFormat = DocumentPageFormat.Jpeg;
                    exporter.CompressionQuality = 90;

                    // export the first page of the document to C:\ on the local system
                    exporter.ExportPage(docInfo, 1, pdfPath);

                    // export all pages to a PDF file, using the above settings (include annotations, and burn-in redactions)

                    // This functionality works great on a server that has .net 4.0.0 installed, there are also some other older windows library dependencies
                    // required for this to work. However, it's great. It has all of the digital annotation stamps included.
                    // The downloaded PDF file looks just like it does in the Laserfiche repository.
                    // I need this same functionality to work in .net core 6.0.
                    exporter.ExportPdf(docInfo, docInfo.AllPages, PdfExportOptions.IncludeText, $"{Directory.GetCurrentDirectory()}\\{fullFilename}");

                    Console.WriteLine($"Finished exporting PDF: {fullFilename}");
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Failed exporting PDF: {fullFilename}");
                Console.WriteLine(ex.Message);
            }
            return fullFilename;
        }

        // Someone suggesting using the Pdf Exporter API in the Pdf Services library but I don't understand
        // how to implement this functionality
        private static string ExportViaPdfExporter(Session session, string filename, int entryId)
        {
            string fullFilename = $"{filename}.pdf";
            try
            {
                using (DocumentInfo docInfo = Document.GetDocumentInfo(entryId, session))
                {
                    // This is the code I found online but I don't see it available in the Pdf Services library.
                    //PdfExporter pdfExporter = new PdfExporter();
                    //// configure the pdfExporter to export images as JPEG, include annotations, and burn-in redactions.
                    //pdfExporter.IncludeAnnotations = true;
                    //pdfExporter.BlackoutRedactions = true;
                    //pdfExporter.PageFormat = DocumentPageFormat.Tiff;
                    //pdfExporter.CompressionQuality = 90;

                    //// export the first page of the document to C:\ on the local system
                    //pdfExporter.ExportPage(docInfo, 1, $"C:\\{filename}.tiff");

                    //// export all pages to a PDF file, using the above settings (include annotations, and burn-in redactions)
                    //pdfExporter.ExportPdf(docInfo, docInfo.AllPages, PdfExportOptions.IncludeText, $"{Directory.GetCurrentDirectory()}\\{fullFilename}");

                    Console.WriteLine($"Finished exporting PDF via Pdf Exporter: {fullFilename}");
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Failed exporting PDF via Pdf Exporter: {fullFilename}");
                Console.WriteLine(ex.Message);
            }
            return fullFilename;
        }

        private static void SaveLaserfichePageAsJpg(string imagePath, Stream stream)
        {
            Console.WriteLine($"Saving pdf image to: {imagePath}...");
            using (Image imageFile = Image.FromStream(stream))
            {
                // The issue with this design is that the saved image does not include the 
                // public digital annotation stamps that exist on the Laserfiche document.
                imageFile.Save(imagePath, ImageFormat.Jpeg);
            }
        }

        private static void AddJpgToPdfPage(PdfDocument pdfCopy, string jpgPath)
        {
            PdfPage page = pdfCopy.AddPage();
            page.Orientation = PageOrientation.Portrait;

            using (XGraphics gfx = XGraphics.FromPdfPage(page))
            {
                XImage image = XImage.FromFile(jpgPath);

                // Draw scaled to fit the pdf page
                int width = (int)(page.Width - 30);
                int height = (int)(page.Height - 30);
                gfx.DrawImage(image, 0, 0, width, height);

                gfx.Dispose();
                image.Dispose();
            }
        }

        private static void DownloadImagePdfDocument(string currentDirectory, PdfDocument pdf, DocumentInfo docInfo)
        {
            using (PageInfoReader pdfReader = docInfo.GetPageInfos())
            {
                foreach (PageInfo PI in pdfReader)
                {
                    using (Stream stream = PI.ReadPagePart(new PagePart()))
                    {
                        if (stream.Length > 0)
                        {
                            string imagePath = $"{currentDirectory}\\{docInfo.Name}({PI.PageNumber}).jpg";
                            File.Delete(imagePath); // Clean up files from previous execution

                            SaveLaserfichePageAsJpg(imagePath, stream);
                            AddJpgToPdfPage(pdf, imagePath);
                        }
                    }
                }
            }
        }

        private static void DownloadElectronicDocument(string currentDirectory, DocumentInfo docInfo)
        {
            string mimeType;
            using (LaserficheReadStream stream = docInfo.ReadEdoc(out mimeType))
            {
                string path = $"{currentDirectory}\\{docInfo.Name}_ElectronicDoc.pdf";
                using (var fileStream = File.Create(path))
                {
                    stream.CopyTo(fileStream);
                }

                if (!File.Exists(path))
                {
                    string msg = $"Failed to write ELECTRONIC Laserfiche PDF document: {docInfo.Id} to: {currentDirectory}";
                    Console.WriteLine(msg);
                }
            }
        }
    }
}
