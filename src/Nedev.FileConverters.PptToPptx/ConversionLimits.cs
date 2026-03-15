namespace Nedev.FileConverters.PptToPptx
{
    /// <summary>
    /// 定义转换操作的限制和约束。
    /// </summary>
    public static class ConversionLimits
    {
        /// <summary>
        /// 默认最大输入文件大小（100 MB）。
        /// </summary>
        public const long DefaultMaxInputFileSize = 100 * 1024 * 1024;

        /// <summary>
        /// 绝对最大输入文件大小（1 GB）。
        /// </summary>
        public const long AbsoluteMaxInputFileSize = 1024 * 1024 * 1024;

        /// <summary>
        /// 默认最大幻灯片数量。
        /// </summary>
        public const int DefaultMaxSlideCount = 1000;

        /// <summary>
        /// 默认最大图片数量。
        /// </summary>
        public const int DefaultMaxImageCount = 5000;

        /// <summary>
        /// 默认最大嵌入资源数量。
        /// </summary>
        public const int DefaultMaxEmbeddedResources = 1000;

        /// <summary>
        /// 单个资源的最大大小（50 MB）。
        /// </summary>
        public const long MaxResourceSize = 50 * 1024 * 1024;

        /// <summary>
        /// 最大递归深度（防止堆栈溢出）。
        /// </summary>
        public const int MaxRecursionDepth = 100;

        /// <summary>
        /// 缓冲区大小（64 KB）。
        /// </summary>
        public const int BufferSize = 64 * 1024;

        /// <summary>
        /// 最大字符串长度（10 MB）。
        /// </summary>
        public const int MaxStringLength = 10 * 1024 * 1024;
    }
}
