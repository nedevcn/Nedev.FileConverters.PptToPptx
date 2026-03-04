namespace Nefdev.PptToPptx
{
    public class PptToPptxConverter
    {
        public static void Convert(string pptPath, string pptxPath)
        {
            using var pptReader = new PptReader(pptPath);
            using var pptxWriter = new PptxWriter(pptxPath);
            
            pptxWriter.WritePresentation(pptReader.ReadPresentation());
        }
    }
}
