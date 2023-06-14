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
        private Session laserficheSession;

        // TODO: Replace with your Laserfiche document that is NOT an electronic document.
        //       This code is for downloading a document comprised of images that have PUBLIC digital annoation stamps applied to them.
        private int entryId = 189271;

        private string currentDirectory = Path.Combine(Directory.GetCurrentDirectory(), "..\\..\\..");

        public DocumentDownload()
        {
            this._laserficheServerName = "Your Laserfiche Config";
            this._laserficheRepoName = "Your Laserfiche Config";
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

                            fileNames.Add(ExportViaCustomCode(this.laserficheSession, "CustomSample", this.entryId));
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
        /// These PDF files are composed of images. There are PUBLIC digital annotation stamps on some of the PDF pages.
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
                        if (!docInfo.IsElectronicDocument)
                        {
                            // This is the code we need help with.
                            // We need to know how to download a PDF document that is comprised of image pages that have public digital annotation stamps
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

                                            Console.WriteLine($"Saving pdf image to: {imagePath}...");
                                            using (Image imageFile = Image.FromStream(stream))
                                            {
                                                // The issue with this design is that the saved image does not include the 
                                                // public digital annotation stamps that exist on the Laserfiche document.
                                                imageFile.Save(imagePath, ImageFormat.Jpeg);
                                            }
                                            AddJpgToPdfPage(pdf, imagePath);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine($"This code doesn't work for electronic documents: {docInfo.Id} to: {pdfPath}");
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
    }
}
