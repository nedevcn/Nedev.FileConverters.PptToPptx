using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace Nedev.FileConverters.PptToPptx.Tests
{
    public class BatchConverterTests
    {
        [Fact]
        public void ConversionPreset_Default_HasExpectedValues()
        {
            var preset = ConversionPreset.Default;
            
            Assert.Equal(ConversionLimits.DefaultMaxInputFileSize, preset.MaxInputFileSize);
            Assert.Equal(ConversionLimits.DefaultMaxSlideCount, preset.MaxSlideCount);
            Assert.Equal(ConversionLimits.DefaultMaxImageCount, preset.MaxImageCount);
            Assert.Equal(TimeSpan.FromMinutes(10), preset.Timeout);
            Assert.False(preset.KeepTempFiles);
        }

        [Fact]
        public void ConversionPreset_Fast_HasStrictLimits()
        {
            var preset = ConversionPreset.Fast;
            
            Assert.Equal(10 * 1024 * 1024, preset.MaxInputFileSize); // 10 MB
            Assert.Equal(100, preset.MaxSlideCount);
            Assert.Equal(500, preset.MaxImageCount);
            Assert.Equal(TimeSpan.FromMinutes(2), preset.Timeout);
            Assert.False(preset.KeepTempFiles);
        }

        [Fact]
        public void ConversionPreset_HighQuality_HasRelaxedLimits()
        {
            var preset = ConversionPreset.HighQuality;
            
            Assert.Equal(500 * 1024 * 1024, preset.MaxInputFileSize); // 500 MB
            Assert.Equal(2000, preset.MaxSlideCount);
            Assert.Equal(10000, preset.MaxImageCount);
            Assert.Equal(TimeSpan.FromMinutes(30), preset.Timeout);
            Assert.True(preset.KeepTempFiles);
        }

        [Fact]
        public void ConversionPreset_Unlimited_HasZeroLimits()
        {
            var preset = ConversionPreset.Unlimited;
            
            Assert.Equal(0, preset.MaxInputFileSize);
            Assert.Equal(0, preset.MaxSlideCount);
            Assert.Equal(0, preset.MaxImageCount);
            Assert.Equal(TimeSpan.Zero, preset.Timeout);
        }

        [Fact]
        public void ConversionPreset_Secure_HasModerateLimits()
        {
            var preset = ConversionPreset.Secure;
            
            Assert.Equal(50 * 1024 * 1024, preset.MaxInputFileSize); // 50 MB
            Assert.Equal(200, preset.MaxSlideCount);
            Assert.Equal(1000, preset.MaxImageCount);
            Assert.Equal(TimeSpan.FromMinutes(5), preset.Timeout);
        }

        [Fact]
        public void ConversionPreset_Batch_HasBalancedLimits()
        {
            var preset = ConversionPreset.Batch;
            
            Assert.Equal(100 * 1024 * 1024, preset.MaxInputFileSize); // 100 MB
            Assert.Equal(500, preset.MaxSlideCount);
            Assert.Equal(2000, preset.MaxImageCount);
            Assert.Equal(TimeSpan.FromMinutes(10), preset.Timeout);
        }

        [Fact]
        public void BatchConversionProgress_Properties_CalculatedCorrectly()
        {
            var results = new List<ConversionResult>
            {
                ConversionResult.SuccessResult(TimeSpan.FromSeconds(1), 1, 1, 0, 100, 80),
                ConversionResult.FailureResult(new Exception("Test"), TimeSpan.FromSeconds(1), inputFileSize: 100)
            };

            var progress = new BatchConversionProgress(2, 5, "test.ppt", results.AsReadOnly());

            Assert.Equal(2, progress.CurrentFileIndex);
            Assert.Equal(5, progress.TotalFiles);
            Assert.Equal("test.ppt", progress.CurrentFilePath);
            Assert.Equal(2, progress.CompletedCount);
            Assert.Equal(1, progress.SuccessCount);
            Assert.Equal(1, progress.FailedCount);
            Assert.Equal(40, progress.OverallPercentComplete); // 2/5 = 40%
        }

        [Fact]
        public void BatchConversionResult_Properties_CalculatedCorrectly()
        {
            var results = new List<ConversionResult>
            {
                ConversionResult.SuccessResult(TimeSpan.FromSeconds(1), 1, 1, 0, 100, 80),
                ConversionResult.SuccessResult(TimeSpan.FromSeconds(1), 1, 1, 0, 100, 80),
                ConversionResult.FailureResult(new Exception("Test"), TimeSpan.FromSeconds(1), inputFileSize: 100),
                ConversionResult.FailureResult(new Exception("Test2"), TimeSpan.FromSeconds(1), inputFileSize: 100)
            };

            var batchResult = new BatchConversionResult(results.AsReadOnly(), TimeSpan.FromSeconds(10));

            Assert.Equal(4, batchResult.TotalCount);
            Assert.Equal(2, batchResult.SuccessCount);
            Assert.Equal(2, batchResult.FailedCount);
            Assert.Equal(50.0, batchResult.SuccessRate); // 2/4 = 50%
            Assert.Equal(TimeSpan.FromSeconds(10), batchResult.TotalDuration);
            Assert.False(batchResult.AllSucceeded);
            Assert.Equal(2, batchResult.GetFailedResults().Count());
            Assert.Equal(2, batchResult.GetSuccessfulResults().Count());
        }

        [Fact]
        public void BatchConversionResult_AllSucceeded_TrueWhenNoFailures()
        {
            var results = new List<ConversionResult>
            {
                ConversionResult.SuccessResult(TimeSpan.FromSeconds(1), 1, 1, 0, 100, 80),
                ConversionResult.SuccessResult(TimeSpan.FromSeconds(1), 1, 1, 0, 100, 80)
            };

            var batchResult = new BatchConversionResult(results.AsReadOnly(), TimeSpan.FromSeconds(5));

            Assert.True(batchResult.AllSucceeded);
            Assert.Equal(100.0, batchResult.SuccessRate);
        }

        [Fact]
        public void BatchConversionResult_EmptyList_HasZeroValues()
        {
            var results = new List<ConversionResult>();
            var batchResult = new BatchConversionResult(results.AsReadOnly(), TimeSpan.Zero);

            Assert.Equal(0, batchResult.TotalCount);
            Assert.Equal(0, batchResult.SuccessCount);
            Assert.Equal(0, batchResult.FailedCount);
            Assert.Equal(0, batchResult.SuccessRate);
            Assert.True(batchResult.AllSucceeded); // No failures = all succeeded
        }
    }
}
