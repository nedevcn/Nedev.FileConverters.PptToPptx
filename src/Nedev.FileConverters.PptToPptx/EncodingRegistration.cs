using System.Text;

namespace Nedev.FileConverters.PptToPptx
{
    internal static class EncodingRegistration
    {
        /// <summary>
        /// Ensure that code page encodings are registered (System.Text.Encoding.CodePages package).
        /// </summary>
        public static void EnsureCodePages()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
    }
}

