using System;
using System.IO;
using Xunit;

namespace Nedev.FileConverters.PptToPptx.Tests
{
    public class ConversionLimitsTests
    {
        [Fact]
        public void ConversionLimits_DefaultValues_AreCorrect()
        {
            Assert.Equal(100 * 1024 * 1024, ConversionLimits.DefaultMaxInputFileSize); // 100 MB
            Assert.Equal(1024L * 1024 * 1024, ConversionLimits.AbsoluteMaxInputFileSize); // 1 GB
            Assert.Equal(1000, ConversionLimits.DefaultMaxSlideCount);
            Assert.Equal(5000, ConversionLimits.DefaultMaxImageCount);
            Assert.Equal(1000, ConversionLimits.DefaultMaxEmbeddedResources);
            Assert.Equal(50 * 1024 * 1024, ConversionLimits.MaxResourceSize); // 50 MB
            Assert.Equal(100, ConversionLimits.MaxRecursionDepth);
            Assert.Equal(64 * 1024, ConversionLimits.BufferSize); // 64 KB
            Assert.Equal(10 * 1024 * 1024, ConversionLimits.MaxStringLength); // 10 MB
        }

        [Fact]
        public void ConversionOptions_DefaultLimits_AreSet()
        {
            var options = new ConversionOptions();
            
            Assert.Equal(ConversionLimits.DefaultMaxInputFileSize, options.MaxInputFileSize);
            Assert.Equal(ConversionLimits.DefaultMaxSlideCount, options.MaxSlideCount);
            Assert.Equal(ConversionLimits.DefaultMaxImageCount, options.MaxImageCount);
        }

        [Fact]
        public void ConversionOptions_CanCustomizeLimits()
        {
            var options = new ConversionOptions
            {
                MaxInputFileSize = 50 * 1024 * 1024, // 50 MB
                MaxSlideCount = 500,
                MaxImageCount = 2000
            };
            
            Assert.Equal(50 * 1024 * 1024, options.MaxInputFileSize);
            Assert.Equal(500, options.MaxSlideCount);
            Assert.Equal(2000, options.MaxImageCount);
        }

        [Fact]
        public void ConversionOptions_ZeroOrNegative_DisablesLimit()
        {
            var options = new ConversionOptions
            {
                MaxInputFileSize = 0,
                MaxSlideCount = -1,
                MaxImageCount = 0
            };
            
            // These should be treated as disabled (0 or negative)
            Assert.Equal(0, options.MaxInputFileSize);
            Assert.Equal(-1, options.MaxSlideCount);
            Assert.Equal(0, options.MaxImageCount);
        }
    }
}
