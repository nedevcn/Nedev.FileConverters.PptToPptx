using System;
using System.IO;
using Xunit;

namespace Nedev.FileConverters.PptToPptx.Tests
{
    public class ExceptionTests
    {
        [Fact]
        public void PptConversionException_DefaultConstructor_SetsMessage()
        {
            var ex = new PptConversionException("Test message");
            Assert.Equal("Test message", ex.Message);
            Assert.Equal(ConversionPhase.Failed, ex.Phase);
        }

        [Fact]
        public void PptConversionException_WithInnerException_SetsInnerException()
        {
            var innerEx = new InvalidOperationException("Inner exception");
            var ex = new PptConversionException("Test message", innerEx);
            Assert.Equal("Test message", ex.Message);
            Assert.Equal(innerEx, ex.InnerException);
        }

        [Fact]
        public void PptConversionException_WithPhase_SetsPhase()
        {
            var ex = new PptConversionException("Test message", ConversionPhase.Reading);
            Assert.Equal(ConversionPhase.Reading, ex.Phase);
        }

        [Fact]
        public void PptConversionException_WithInputPath_SetsInputPath()
        {
            var ex = new PptConversionException("Test message", ConversionPhase.Reading, "test.ppt");
            Assert.Equal("test.ppt", ex.InputPath);
        }

        [Fact]
        public void InvalidPptFormatException_DefaultConstructor_SetsPhaseToReading()
        {
            var ex = new InvalidPptFormatException("Invalid format");
            Assert.Equal("Invalid format", ex.Message);
            Assert.Equal(ConversionPhase.Reading, ex.Phase);
        }

        [Fact]
        public void InvalidPptFormatException_WithInnerException_SetsInnerException()
        {
            var innerEx = new InvalidDataException("Data error");
            var ex = new InvalidPptFormatException("Invalid format", innerEx);
            Assert.Equal(innerEx, ex.InnerException);
        }

        [Fact]
        public void OleCompoundFileException_DefaultConstructor_SetsPhaseToReading()
        {
            var ex = new OleCompoundFileException("OLE error");
            Assert.Equal("OLE error", ex.Message);
            Assert.Equal(ConversionPhase.Reading, ex.Phase);
        }

        [Fact]
        public void ConversionCanceledException_DefaultConstructor_SetsMessage()
        {
            var ex = new ConversionCanceledException();
            // Message is in Chinese "转换操作已被取消。"
            Assert.Contains("取消", ex.Message);
            Assert.Equal(ConversionPhase.Failed, ex.Phase);
        }

        [Fact]
        public void ConversionCanceledException_WithCustomMessage_SetsMessage()
        {
            var ex = new ConversionCanceledException("Custom cancel message");
            Assert.Equal("Custom cancel message", ex.Message);
        }
    }
}
