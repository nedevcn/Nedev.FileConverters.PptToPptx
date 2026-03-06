using System;
using Nedev.FileConverters.PptToPptx;
using Nedev.FileConverters.Core;

namespace Nedev.FileConverters.PptToPptx.Cli
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.Error.WriteLine("Usage: ppt2pptx <input.ppt> <output.pptx>");
                return 1;
            }

            var input = args[0];
            var output = args[1];

            try
            {
                Console.WriteLine($"Converting '{input}' -> '{output}'...");
                PptToPptxConverter.Convert(input, output);
                Console.WriteLine("Conversion succeeded.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Conversion failed: " + ex.Message);
                return 1;
            }

            return 0;
        }
    }
}