using System;
using System.IO;
using Nedev.FileConverters.Core;

namespace Nedev.FileConverters.PptToPptx
{
    /// <summary>
    /// Adapter that allows the PPT→PPTX functionality to participate in the
    /// shared converter registry provided by the NuGet package.  The core
    /// infrastructure automatically discovers any class marked with
    /// <see cref="FileConverterAttribute"/>.
    /// </summary>
    [FileConverter("ppt", "pptx")]
    public class PptToPptxFileConverter : IFileConverter
    {
        public Stream Convert(Stream input)
        {
            // write input stream to temporary file
            string tempIn = Path.GetTempFileName();
            string tempOut = Path.GetTempFileName();
            try
            {
                using (var fs = File.OpenWrite(tempIn))
                {
                    input.CopyTo(fs);
                }

                PptToPptxConverter.Convert(tempIn, tempOut);
                var outputStream = new MemoryStream();
                using (var fs = File.OpenRead(tempOut))
                {
                    fs.CopyTo(outputStream);
                }
                outputStream.Position = 0;
                return outputStream;
            }
            finally
            {
                try { File.Delete(tempIn); } catch { }
                try { File.Delete(tempOut); } catch { }
            }
        }
    }
}