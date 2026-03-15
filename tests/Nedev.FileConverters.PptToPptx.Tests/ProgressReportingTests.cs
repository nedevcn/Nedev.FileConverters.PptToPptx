using System;
using System.Collections.Generic;
using System.Threading;
using Xunit;

namespace Nedev.FileConverters.PptToPptx.Tests
{
    public class ProgressReportingTests
    {
        [Fact]
        public void ConversionProgress_Constructor_SetsProperties()
        {
            var progress = new ConversionProgress(ConversionPhase.Reading, 50, "Test message", 5, 10);
            
            Assert.Equal(ConversionPhase.Reading, progress.Phase);
            Assert.Equal(50, progress.PercentComplete);
            Assert.Equal("Test message", progress.Message);
            Assert.Equal(5, progress.SlidesProcessed);
            Assert.Equal(10, progress.TotalSlides);
        }

        [Fact]
        public void ConversionProgress_PercentComplete_IsClampedTo100()
        {
            var progress = new ConversionProgress(ConversionPhase.Reading, 150, "Test");
            Assert.Equal(100, progress.PercentComplete);
        }

        [Fact]
        public void ConversionProgress_PercentComplete_IsClampedTo0()
        {
            var progress = new ConversionProgress(ConversionPhase.Reading, -10, "Test");
            Assert.Equal(0, progress.PercentComplete);
        }

        [Fact]
        public void ConversionProgress_NullMessage_IsEmptyString()
        {
            var progress = new ConversionProgress(ConversionPhase.Reading, 50, null!);
            Assert.Equal(string.Empty, progress.Message);
        }

        [Fact]
        public void ConversionOptions_ReportProgress_InvokesCallback()
        {
            var progressReports = new List<ConversionProgress>();
            var options = new ConversionOptions
            {
                Progress = progress => progressReports.Add(progress)
            };

            options.ReportProgress(ConversionPhase.Initializing, 0, "Starting");
            options.ReportProgress(ConversionPhase.Reading, 50, "Reading");
            options.ReportProgress(ConversionPhase.Completed, 100, "Done", 5, 5);

            Assert.Equal(3, progressReports.Count);
            Assert.Equal(ConversionPhase.Initializing, progressReports[0].Phase);
            Assert.Equal(0, progressReports[0].PercentComplete);
            Assert.Equal(ConversionPhase.Reading, progressReports[1].Phase);
            Assert.Equal(50, progressReports[1].PercentComplete);
            Assert.Equal(ConversionPhase.Completed, progressReports[2].Phase);
            Assert.Equal(100, progressReports[2].PercentComplete);
            Assert.Equal(5, progressReports[2].SlidesProcessed);
            Assert.Equal(5, progressReports[2].TotalSlides);
        }

        [Fact]
        public void ConversionOptions_ReportProgress_NoCallback_DoesNotThrow()
        {
            var options = new ConversionOptions();
            
            // Should not throw even though Progress is null
            options.ReportProgress(ConversionPhase.Reading, 50, "Test");
        }

        [Fact]
        public void ConversionOptions_LogMessage_InvokesCallback()
        {
            var logMessages = new List<string>();
            var options = new ConversionOptions
            {
                Log = message => logMessages.Add(message)
            };

            options.LogMessage("Test message 1");
            options.LogMessage("Test message 2");

            Assert.Equal(2, logMessages.Count);
            Assert.Equal("Test message 1", logMessages[0]);
            Assert.Equal("Test message 2", logMessages[1]);
        }

        [Fact]
        public void ConversionOptions_LogMessage_NoCallback_DoesNotThrow()
        {
            var options = new ConversionOptions();
            
            // Should not throw even though Log is null
            options.LogMessage("Test message");
        }

        [Fact]
        public void ConversionOptions_DefaultValues_AreCorrect()
        {
            var options = new ConversionOptions();
            
            Assert.Null(options.Log);
            Assert.Null(options.Progress);
            Assert.False(options.KeepTempFiles);
            Assert.Null(options.PptAnsiCodePageOverride);
            Assert.Null(options.BiffAnsiCodePageOverride);
        }

        [Theory]
        [InlineData(ConversionPhase.Initializing)]
        [InlineData(ConversionPhase.Reading)]
        [InlineData(ConversionPhase.ProcessingStructure)]
        [InlineData(ConversionPhase.ExtractingMedia)]
        [InlineData(ConversionPhase.ProcessingSlides)]
        [InlineData(ConversionPhase.Writing)]
        [InlineData(ConversionPhase.Finalizing)]
        [InlineData(ConversionPhase.Completed)]
        [InlineData(ConversionPhase.Failed)]
        public void ConversionPhase_AllPhases_AreDefined(ConversionPhase phase)
        {
            // This test ensures all expected phases are defined in the enum
            Assert.True(Enum.IsDefined(typeof(ConversionPhase), phase));
        }
    }
}
