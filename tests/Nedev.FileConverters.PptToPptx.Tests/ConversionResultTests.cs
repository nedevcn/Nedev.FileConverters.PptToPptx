using System;
using System.IO;
using Xunit;

namespace Nedev.FileConverters.PptToPptx.Tests
{
    public class ConversionResultTests
    {
        [Fact]
        public void ConversionResult_SuccessResult_HasCorrectProperties()
        {
            var duration = TimeSpan.FromSeconds(5);
            var result = ConversionResult.SuccessResult(
                duration,
                slideCount: 10,
                imageCount: 5,
                embeddedResourceCount: 2,
                inputFileSize: 1000,
                outputFileSize: 800);

            Assert.True(result.Success);
            Assert.Null(result.Exception);
            Assert.Equal(duration, result.Duration);
            Assert.Equal(10, result.SlideCount);
            Assert.Equal(5, result.ImageCount);
            Assert.Equal(2, result.EmbeddedResourceCount);
            Assert.Equal(1000, result.InputFileSize);
            Assert.Equal(800, result.OutputFileSize);
            Assert.True(result.CompletedAt <= DateTime.UtcNow);
        }

        [Fact]
        public void ConversionResult_FailureResult_HasCorrectProperties()
        {
            var duration = TimeSpan.FromSeconds(3);
            var exception = new InvalidOperationException("Test error");
            
            var result = ConversionResult.FailureResult(
                exception,
                duration,
                slideCount: 5,
                imageCount: 2,
                embeddedResourceCount: 1,
                inputFileSize: 500,
                outputFileSize: 0);

            Assert.False(result.Success);
            Assert.Same(exception, result.Exception);
            Assert.Equal(duration, result.Duration);
            Assert.Equal(5, result.SlideCount);
            Assert.Equal(2, result.ImageCount);
            Assert.Equal(1, result.EmbeddedResourceCount);
            Assert.Equal(500, result.InputFileSize);
            Assert.Equal(0, result.OutputFileSize);
        }

        [Fact]
        public void ConversionResult_CompressionRatio_CalculatedCorrectly()
        {
            var result = ConversionResult.SuccessResult(
                TimeSpan.FromSeconds(1),
                slideCount: 1,
                imageCount: 1,
                embeddedResourceCount: 0,
                inputFileSize: 1000,
                outputFileSize: 800);

            Assert.Equal(0.8, result.CompressionRatio);
        }

        [Fact]
        public void ConversionResult_CompressionRatio_ZeroWhenNoInput()
        {
            var result = ConversionResult.SuccessResult(
                TimeSpan.FromSeconds(1),
                slideCount: 1,
                imageCount: 1,
                embeddedResourceCount: 0,
                inputFileSize: 0,
                outputFileSize: 800);

            Assert.Equal(0, result.CompressionRatio);
        }

        [Fact]
        public void ConversionResult_SlidesPerSecond_CalculatedCorrectly()
        {
            var result = ConversionResult.SuccessResult(
                TimeSpan.FromSeconds(10),
                slideCount: 20,
                imageCount: 5,
                embeddedResourceCount: 0,
                inputFileSize: 1000,
                outputFileSize: 800);

            Assert.Equal(2.0, result.SlidesPerSecond);
        }

        [Fact]
        public void ConversionResult_SlidesPerSecond_ZeroWhenNoDuration()
        {
            var result = ConversionResult.SuccessResult(
                TimeSpan.Zero,
                slideCount: 10,
                imageCount: 5,
                embeddedResourceCount: 0,
                inputFileSize: 1000,
                outputFileSize: 800);

            Assert.Equal(0, result.SlidesPerSecond);
        }

        [Fact]
        public void ConversionOptions_DefaultTimeout_IsTenMinutes()
        {
            var options = new ConversionOptions();
            Assert.Equal(TimeSpan.FromMinutes(10), options.Timeout);
        }

        [Fact]
        public void ConversionOptions_CanCustomizeTimeout()
        {
            var options = new ConversionOptions
            {
                Timeout = TimeSpan.FromMinutes(5)
            };
            Assert.Equal(TimeSpan.FromMinutes(5), options.Timeout);
        }

        [Fact]
        public void ConversionOptions_ZeroTimeout_DisablesTimeout()
        {
            var options = new ConversionOptions
            {
                Timeout = TimeSpan.Zero
            };
            Assert.Equal(TimeSpan.Zero, options.Timeout);
        }
    }
}
