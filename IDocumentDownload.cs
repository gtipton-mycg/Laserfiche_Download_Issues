namespace Laserfiche_Download_Issues
{
    public interface IDocumentDownload
    {
        IEnumerable<string> DownloadLaserficheDocument();
    }
}