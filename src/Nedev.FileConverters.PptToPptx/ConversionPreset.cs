using System;

namespace Nedev.FileConverters.PptToPptx
{
    /// <summary>
    /// 预定义的转换配置预设。
    /// </summary>
    public static class ConversionPreset
    {
        /// <summary>
        /// 默认预设 - 平衡性能和质量。
        /// </summary>
        public static ConversionOptions Default => new ConversionOptions();

        /// <summary>
        /// 快速预设 - 最小资源使用，适合简单文档。
        /// </summary>
        public static ConversionOptions Fast => new ConversionOptions
        {
            MaxInputFileSize = 10 * 1024 * 1024,  // 10 MB
            MaxSlideCount = 100,
            MaxImageCount = 500,
            Timeout = TimeSpan.FromMinutes(2),
            KeepTempFiles = false
        };

        /// <summary>
        /// 高质量预设 - 保留更多细节，适合复杂文档。
        /// </summary>
        public static ConversionOptions HighQuality => new ConversionOptions
        {
            MaxInputFileSize = 500 * 1024 * 1024,  // 500 MB
            MaxSlideCount = 2000,
            MaxImageCount = 10000,
            Timeout = TimeSpan.FromMinutes(30),
            KeepTempFiles = true  // 保留临时文件以便调试
        };

        /// <summary>
        /// 无限制预设 - 最小限制，适合服务器环境（谨慎使用）。
        /// </summary>
        public static ConversionOptions Unlimited => new ConversionOptions
        {
            MaxInputFileSize = 0,  // 禁用限制
            MaxSlideCount = 0,
            MaxImageCount = 0,
            Timeout = TimeSpan.Zero,  // 禁用超时
            KeepTempFiles = false
        };

        /// <summary>
        /// 安全预设 - 严格限制，适合不受信任的来源。
        /// </summary>
        public static ConversionOptions Secure => new ConversionOptions
        {
            MaxInputFileSize = 50 * 1024 * 1024,  // 50 MB
            MaxSlideCount = 200,
            MaxImageCount = 1000,
            Timeout = TimeSpan.FromMinutes(5),
            KeepTempFiles = false
        };

        /// <summary>
        /// 批量处理预设 - 优化多文件处理。
        /// </summary>
        public static ConversionOptions Batch => new ConversionOptions
        {
            MaxInputFileSize = 100 * 1024 * 1024,  // 100 MB
            MaxSlideCount = 500,
            MaxImageCount = 2000,
            Timeout = TimeSpan.FromMinutes(10),
            KeepTempFiles = false
        };
    }
}
