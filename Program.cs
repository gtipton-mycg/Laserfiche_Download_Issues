// See https://aka.ms/new-console-template for more information

namespace Laserfiche_Download_Issues
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("*** User Notes ***");
            Console.WriteLine("Demonstrate issues with the download functionality in the Laserfiche SDK.");
            Console.WriteLine("The issue we have found is when downloading TIFF PDF files that contain digital annotation stamps.");
            Console.WriteLine("These digital stamps are public images that are used to overlay text and signatures on the PDF files.");
            Console.WriteLine("These stamps are applied by Laserfiche repository users.");
            Console.WriteLine("*** End Of User Notes ***");
            Console.WriteLine("--");

            Console.WriteLine("Calling the Laserfiche Download sample code");
            IDocumentDownload documentDownload = new DocumentDownload();
            documentDownload.DownloadLaserficheDocument();
            Console.WriteLine("Finished calling the Laserfiche Download sample code");
            Console.WriteLine("Please close this window or click ctrl + c");

            // Pause the Console so we can read the output.
            // Click ctrl + c or close the open dialog to exit the program
            Console.ReadLine(); // pause after completion
        }
    }
}