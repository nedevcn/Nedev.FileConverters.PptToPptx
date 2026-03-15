using System;

namespace Nedev.FileConverters.PptToPptx
{
    /// <summary>
    /// 表示 PPT 到 PPTX 转换操作的结果。
    /// </summary>
    public class ConversionResult
    {
        /// <summary>
        /// 获取一个值，指示转换是否成功。
        /// </summary>
        public bool Success { get; }

        /// <summary>
        /// 获取转换过程中发生的异常（如果有）。
        /// </summary>
        public Exception? Exception { get; }

        /// <summary>
        /// 获取转换的持续时间。
        /// </summary>
        public TimeSpan Duration { get; }

        /// <summary>
        /// 获取处理的幻灯片数量。
        /// </summary>
        public int SlideCount { get; }

        /// <summary>
        /// 获取处理的图片数量。
        /// </summary>
        public int ImageCount { get; }

        /// <summary>
        /// 获取处理的嵌入资源数量。
        /// </summary>
        public int EmbeddedResourceCount { get; }

        /// <summary>
        /// 获取输出文件大小（字节）。
        /// </summary>
        public long OutputFileSize { get; }

        /// <summary>
        /// 获取输入文件大小（字节）。
        /// </summary>
        public long InputFileSize { get; }

        /// <summary>
        /// 获取转换完成的时间戳。
        /// </summary>
        public DateTime CompletedAt { get; }

        /// <summary>
        /// 初始化 <see cref="ConversionResult"/> 类的新实例。
        /// </summary>
        public ConversionResult(
            bool success,
            Exception? exception,
            TimeSpan duration,
            int slideCount,
            int imageCount,
            int embeddedResourceCount,
            long inputFileSize,
            long outputFileSize)
        {
            Success = success;
            Exception = exception;
            Duration = duration;
            SlideCount = slideCount;
            ImageCount = imageCount;
            EmbeddedResourceCount = embeddedResourceCount;
            InputFileSize = inputFileSize;
            OutputFileSize = outputFileSize;
            CompletedAt = DateTime.UtcNow;
        }

        /// <summary>
        /// 创建一个表示成功转换的结果。
        /// </summary>
        public static ConversionResult SuccessResult(
            TimeSpan duration,
            int slideCount,
            int imageCount,
            int embeddedResourceCount,
            long inputFileSize,
            long outputFileSize)
        {
            return new ConversionResult(true, null, duration, slideCount, imageCount, embeddedResourceCount, inputFileSize, outputFileSize);
        }

        /// <summary>
        /// 创建一个表示失败转换的结果。
        /// </summary>
        public static ConversionResult FailureResult(
            Exception exception,
            TimeSpan duration,
            int slideCount = 0,
            int imageCount = 0,
            int embeddedResourceCount = 0,
            long inputFileSize = 0,
            long outputFileSize = 0)
        {
            return new ConversionResult(false, exception, duration, slideCount, imageCount, embeddedResourceCount, inputFileSize, outputFileSize);
        }

        /// <summary>
        /// 获取转换的压缩比率（输出大小 / 输入大小）。
        /// </summary>
        public double CompressionRatio => InputFileSize > 0 ? (double)OutputFileSize / InputFileSize : 0;

        /// <summary>
        /// 获取每秒处理的幻灯片数量。
        /// </summary>
        public double SlidesPerSecond => Duration.TotalSeconds > 0 ? SlideCount / Duration.TotalSeconds : 0;
    }
}
