using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Nedev.FileConverters.PptToPptx
{
    /// <summary>
    /// 表示批处理转换操作的进度信息。
    /// </summary>
    public class BatchConversionProgress
    {
        /// <summary>
        /// 获取当前处理的文件索引（从1开始）。
        /// </summary>
        public int CurrentFileIndex { get; }

        /// <summary>
        /// 获取总文件数量。
        /// </summary>
        public int TotalFiles { get; }

        /// <summary>
        /// 获取当前处理的文件路径。
        /// </summary>
        public string CurrentFilePath { get; }

        /// <summary>
        /// 获取已完成的转换结果列表。
        /// </summary>
        public IReadOnlyList<ConversionResult> CompletedResults { get; }

        /// <summary>
        /// 获取已处理的文件数量。
        /// </summary>
        public int CompletedCount => CompletedResults.Count;

        /// <summary>
        /// 获取成功的转换数量。
        /// </summary>
        public int SuccessCount => CompletedResults.Count(r => r.Success);

        /// <summary>
        /// 获取失败的转换数量。
        /// </summary>
        public int FailedCount => CompletedResults.Count(r => !r.Success);

        /// <summary>
        /// 获取总体完成百分比。
        /// </summary>
        public int OverallPercentComplete => TotalFiles > 0 ? (CompletedCount * 100) / TotalFiles : 0;

        /// <summary>
        /// 初始化 <see cref="BatchConversionProgress"/> 类的新实例。
        /// </summary>
        public BatchConversionProgress(int currentFileIndex, int totalFiles, string currentFilePath, IReadOnlyList<ConversionResult> completedResults)
        {
            CurrentFileIndex = currentFileIndex;
            TotalFiles = totalFiles;
            CurrentFilePath = currentFilePath;
            CompletedResults = completedResults;
        }
    }

    /// <summary>
    /// 表示批处理转换的结果。
    /// </summary>
    public class BatchConversionResult
    {
        /// <summary>
        /// 获取所有转换结果。
        /// </summary>
        public IReadOnlyList<ConversionResult> Results { get; }

        /// <summary>
        /// 获取成功的转换数量。
        /// </summary>
        public int SuccessCount => Results.Count(r => r.Success);

        /// <summary>
        /// 获取失败的转换数量。
        /// </summary>
        public int FailedCount => Results.Count(r => !r.Success);

        /// <summary>
        /// 获取总文件数量。
        /// </summary>
        public int TotalCount => Results.Count;

        /// <summary>
        /// 获取成功率百分比。
        /// </summary>
        public double SuccessRate => TotalCount > 0 ? (double)SuccessCount / TotalCount * 100 : 0;

        /// <summary>
        /// 获取总耗时。
        /// </summary>
        public TimeSpan TotalDuration { get; }

        /// <summary>
        /// 获取一个值，指示是否所有转换都成功。
        /// </summary>
        public bool AllSucceeded => FailedCount == 0;

        /// <summary>
        /// 初始化 <see cref="BatchConversionResult"/> 类的新实例。
        /// </summary>
        public BatchConversionResult(IReadOnlyList<ConversionResult> results, TimeSpan totalDuration)
        {
            Results = results;
            TotalDuration = totalDuration;
        }

        /// <summary>
        /// 获取失败的转换结果。
        /// </summary>
        public IEnumerable<ConversionResult> GetFailedResults() => Results.Where(r => !r.Success);

        /// <summary>
        /// 获取成功的转换结果。
        /// </summary>
        public IEnumerable<ConversionResult> GetSuccessfulResults() => Results.Where(r => r.Success);
    }

    /// <summary>
    /// 提供批量 PPT 到 PPTX 转换功能。
    /// </summary>
    public static class BatchConverter
    {
        /// <summary>
        /// 批量转换多个 PPT 文件。
        /// </summary>
        /// <param name="inputOutputPairs">输入输出文件路径对的集合。</param>
        /// <param name="options">转换选项。</param>
        /// <param name="progress">进度回调。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <returns>批处理转换结果。</returns>
        public static async Task<BatchConversionResult> ConvertAsync(
            IEnumerable<(string inputPath, string outputPath)> inputOutputPairs,
            ConversionOptions? options = null,
            Action<BatchConversionProgress>? progress = null,
            CancellationToken cancellationToken = default)
        {
            var pairs = inputOutputPairs.ToList();
            var results = new List<ConversionResult>();
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            for (int i = 0; i < pairs.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var (inputPath, outputPath) = pairs[i];
                
                // 报告进度
                var batchProgress = new BatchConversionProgress(
                    i + 1,
                    pairs.Count,
                    inputPath,
                    results.AsReadOnly());
                progress?.Invoke(batchProgress);

                // 执行转换
                var result = await PptToPptxConverter.ConvertWithResultAsync(
                    inputPath,
                    outputPath,
                    options,
                    cancellationToken);
                
                results.Add(result);
            }

            stopwatch.Stop();
            return new BatchConversionResult(results.AsReadOnly(), stopwatch.Elapsed);
        }

        /// <summary>
        /// 批量转换目录中的所有 PPT 文件。
        /// </summary>
        /// <param name="inputDirectory">输入目录路径。</param>
        /// <param name="outputDirectory">输出目录路径。</param>
        /// <param name="options">转换选项。</param>
        /// <param name="progress">进度回调。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <param name="searchPattern">搜索模式（默认：*.ppt）。</param>
        /// <returns>批处理转换结果。</returns>
        public static Task<BatchConversionResult> ConvertDirectoryAsync(
            string inputDirectory,
            string outputDirectory,
            ConversionOptions? options = null,
            Action<BatchConversionProgress>? progress = null,
            CancellationToken cancellationToken = default,
            string searchPattern = "*.ppt")
        {
            if (!Directory.Exists(inputDirectory))
                throw new DirectoryNotFoundException($"Input directory not found: {inputDirectory}");

            Directory.CreateDirectory(outputDirectory);

            var inputFiles = Directory.GetFiles(inputDirectory, searchPattern, SearchOption.TopDirectoryOnly);
            
            var pairs = inputFiles.Select(inputPath =>
            {
                var fileName = Path.GetFileNameWithoutExtension(inputPath) + ".pptx";
                var outputPath = Path.Combine(outputDirectory, fileName);
                return (inputPath, outputPath);
            });

            return ConvertAsync(pairs, options, progress, cancellationToken);
        }

        /// <summary>
        /// 并行批量转换多个 PPT 文件。
        /// </summary>
        /// <param name="inputOutputPairs">输入输出文件路径对的集合。</param>
        /// <param name="options">转换选项。</param>
        /// <param name="maxParallelism">最大并行度（默认：处理器数量）。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <returns>批处理转换结果。</returns>
        public static async Task<BatchConversionResult> ConvertParallelAsync(
            IEnumerable<(string inputPath, string outputPath)> inputOutputPairs,
            ConversionOptions? options = null,
            int? maxParallelism = null,
            CancellationToken cancellationToken = default)
        {
            var pairs = inputOutputPairs.ToList();
            var results = new List<ConversionResult>(new ConversionResult[pairs.Count]);
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = maxParallelism ?? Environment.ProcessorCount,
                CancellationToken = cancellationToken
            };

            await Task.Run(() =>
            {
                Parallel.For(0, pairs.Count, parallelOptions, i =>
                {
                    var (inputPath, outputPath) = pairs[i];
                    
                    try
                    {
                        var result = PptToPptxConverter.ConvertWithResult(
                            inputPath,
                            outputPath,
                            options,
                            cancellationToken);
                        results[i] = result;
                    }
                    catch (Exception ex)
                    {
                        results[i] = ConversionResult.FailureResult(
                            ex,
                            TimeSpan.Zero,
                            inputFileSize: new FileInfo(inputPath).Length);
                    }
                });
            }, cancellationToken);

            stopwatch.Stop();
            return new BatchConversionResult(results.AsReadOnly(), stopwatch.Elapsed);
        }
    }
}
