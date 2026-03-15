using System;

namespace Nedev.FileConverters.PptToPptx
{
    /// <summary>
    /// 表示 PPT 转换过程中发生的错误。
    /// </summary>
    public class PptConversionException : Exception
    {
        /// <summary>
        /// 获取转换失败的阶段。
        /// </summary>
        public ConversionPhase Phase { get; }

        /// <summary>
        /// 获取输入文件路径（如果可用）。
        /// </summary>
        public string? InputPath { get; }

        /// <summary>
        /// 初始化 <see cref="PptConversionException"/> 类的新实例。
        /// </summary>
        public PptConversionException(string message)
            : base(message)
        {
            Phase = ConversionPhase.Failed;
        }

        /// <summary>
        /// 初始化 <see cref="PptConversionException"/> 类的新实例。
        /// </summary>
        public PptConversionException(string message, Exception innerException)
            : base(message, innerException)
        {
            Phase = ConversionPhase.Failed;
        }

        /// <summary>
        /// 初始化 <see cref="PptConversionException"/> 类的新实例。
        /// </summary>
        public PptConversionException(string message, ConversionPhase phase, string? inputPath = null)
            : base(message)
        {
            Phase = phase;
            InputPath = inputPath;
        }

        /// <summary>
        /// 初始化 <see cref="PptConversionException"/> 类的新实例。
        /// </summary>
        public PptConversionException(string message, ConversionPhase phase, Exception innerException, string? inputPath = null)
            : base(message, innerException)
        {
            Phase = phase;
            InputPath = inputPath;
        }
    }

    /// <summary>
    /// 表示 PPT 文件格式无效或损坏时发生的错误。
    /// </summary>
    public class InvalidPptFormatException : PptConversionException
    {
        /// <summary>
        /// 初始化 <see cref="InvalidPptFormatException"/> 类的新实例。
        /// </summary>
        public InvalidPptFormatException(string message)
            : base(message, ConversionPhase.Reading)
        {
        }

        /// <summary>
        /// 初始化 <see cref="InvalidPptFormatException"/> 类的新实例。
        /// </summary>
        public InvalidPptFormatException(string message, Exception innerException)
            : base(message, ConversionPhase.Reading, innerException)
        {
        }
    }

    /// <summary>
    /// 表示 OLE 复合文件解析失败时发生的错误。
    /// </summary>
    public class OleCompoundFileException : InvalidPptFormatException
    {
        /// <summary>
        /// 初始化 <see cref="OleCompoundFileException"/> 类的新实例。
        /// </summary>
        public OleCompoundFileException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// 初始化 <see cref="OleCompoundFileException"/> 类的新实例。
        /// </summary>
        public OleCompoundFileException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// 表示转换操作被取消时发生的错误。
    /// </summary>
    public class ConversionCanceledException : PptConversionException
    {
        /// <summary>
        /// 初始化 <see cref="ConversionCanceledException"/> 类的新实例。
        /// </summary>
        public ConversionCanceledException()
            : base("转换操作已被取消。", ConversionPhase.Failed)
        {
        }

        /// <summary>
        /// 初始化 <see cref="ConversionCanceledException"/> 类的新实例。
        /// </summary>
        public ConversionCanceledException(string message)
            : base(message, ConversionPhase.Failed)
        {
        }
    }
}
