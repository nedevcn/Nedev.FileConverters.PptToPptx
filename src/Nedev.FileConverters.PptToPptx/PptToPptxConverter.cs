using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Nedev.FileConverters.PptToPptx
{
    /// <summary>
    /// Provides methods for converting legacy PowerPoint (.ppt) files to modern OpenXML (.pptx) format.
    /// </summary>
    public static class PptToPptxConverter
    {
        /// <summary>
        /// Converts a .ppt file to .pptx format.
        /// </summary>
        /// <param name="pptPath">The path to the input .ppt file.</param>
        /// <param name="pptxPath">The path for the output .pptx file.</param>
        /// <exception cref="ArgumentException">Thrown when path arguments are invalid.</exception>
        /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
        public static void Convert(string pptPath, string pptxPath)
        {
            Convert(pptPath, pptxPath, null, CancellationToken.None);
        }

        /// <summary>
        /// Converts a .ppt file to .pptx format with optional configuration.
        /// </summary>
        /// <param name="pptPath">The path to the input .ppt file.</param>
        /// <param name="pptxPath">The path for the output .pptx file.</param>
        /// <param name="options">Optional conversion options and callbacks.</param>
        /// <exception cref="ArgumentException">Thrown when path arguments are invalid.</exception>
        /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
        public static void Convert(string pptPath, string pptxPath, ConversionOptions? options)
        {
            Convert(pptPath, pptxPath, options, CancellationToken.None);
        }

        /// <summary>
        /// Converts a .ppt file to .pptx format with optional configuration and cancellation support.
        /// </summary>
        /// <param name="pptPath">The path to the input .ppt file.</param>
        /// <param name="pptxPath">The path for the output .pptx file.</param>
        /// <param name="options">Optional conversion options and callbacks.</param>
        /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
        /// <exception cref="ArgumentException">Thrown when path arguments are invalid.</exception>
        /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
        /// <exception cref="OperationCanceledException">Thrown when the operation is canceled.</exception>
        public static void Convert(string pptPath, string pptxPath, ConversionOptions? options, CancellationToken cancellationToken)
        {
            ValidatePaths(pptPath, pptxPath, options);

            var outDir = Path.GetDirectoryName(pptxPath);
            if (!string.IsNullOrEmpty(outDir))
                Directory.CreateDirectory(outDir);

            cancellationToken.ThrowIfCancellationRequested();
            options?.ReportProgress(ConversionPhase.Initializing, 0, "Starting conversion...");

            try
            {
                Presentation presentation;

                // Read phase
                cancellationToken.ThrowIfCancellationRequested();
                options?.ReportProgress(ConversionPhase.Reading, 10, "Reading PPT file...");
                using (var pptReader = new PptReader(pptPath, options))
                {
                    presentation = pptReader.ReadPresentation(cancellationToken);
                }

                // Validate presentation limits
                ValidatePresentation(presentation, options);

                cancellationToken.ThrowIfCancellationRequested();
                options?.ReportProgress(ConversionPhase.ProcessingStructure, 30, $"Found {presentation.Slides.Count} slides...", 0, presentation.Slides.Count);

                // Write phase
                cancellationToken.ThrowIfCancellationRequested();
                options?.ReportProgress(ConversionPhase.Writing, 50, "Writing PPTX file...");
                using (var pptxWriter = new PptxWriter(pptxPath, options))
                {
                    pptxWriter.WritePresentation(presentation, cancellationToken);
                }

                cancellationToken.ThrowIfCancellationRequested();
                options?.ReportProgress(ConversionPhase.Finalizing, 90, "Finalizing...");
                options?.ReportProgress(ConversionPhase.Completed, 100, "Conversion completed successfully.", presentation.Slides.Count, presentation.Slides.Count);
            }
            catch (OperationCanceledException)
            {
                options?.ReportProgress(ConversionPhase.Failed, 0, "Conversion canceled.");
                throw;
            }
            catch (PptConversionException)
            {
                options?.ReportProgress(ConversionPhase.Failed, 0, "Conversion failed.");
                throw;
            }
            catch (Exception ex)
            {
                options?.ReportProgress(ConversionPhase.Failed, 0, $"Conversion failed: {ex.Message}");
                throw new PptConversionException($"Failed to convert '{pptPath}' to '{pptxPath}'.", ConversionPhase.Failed, ex, pptPath);
            }
        }

        /// <summary>
        /// Converts a .ppt file to .pptx format and returns detailed result information.
        /// </summary>
        /// <param name="pptPath">The path to the input .ppt file.</param>
        /// <param name="pptxPath">The path for the output .pptx file.</param>
        /// <param name="options">Optional conversion options and callbacks.</param>
        /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
        /// <returns>A <see cref="ConversionResult"/> containing conversion details.</returns>
        public static ConversionResult ConvertWithResult(string pptPath, string pptxPath, ConversionOptions? options = null, CancellationToken cancellationToken = default)
        {
            var stopwatch = Stopwatch.StartNew();
            var inputFileSize = new FileInfo(pptPath).Length;
            
            try
            {
                Convert(pptPath, pptxPath, options, cancellationToken);
                
                stopwatch.Stop();
                var outputFileSize = new FileInfo(pptxPath).Length;
                
                // Note: These counts would need to be tracked during conversion
                // For now, return basic information
                return ConversionResult.SuccessResult(
                    stopwatch.Elapsed,
                    slideCount: 0,  // Would need to track during conversion
                    imageCount: 0,  // Would need to track during conversion
                    embeddedResourceCount: 0,  // Would need to track during conversion
                    inputFileSize,
                    outputFileSize);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                return ConversionResult.FailureResult(ex, stopwatch.Elapsed, inputFileSize: inputFileSize);
            }
        }

        /// <summary>
        /// Asynchronously converts a .ppt file to .pptx format.
        /// </summary>
        /// <param name="pptPath">The path to the input .ppt file.</param>
        /// <param name="pptxPath">The path for the output .pptx file.</param>
        /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
        /// <returns>A task representing the asynchronous conversion operation.</returns>
        public static Task ConvertAsync(string pptPath, string pptxPath, CancellationToken cancellationToken = default)
        {
            return ConvertAsync(pptPath, pptxPath, null, cancellationToken);
        }

        /// <summary>
        /// Asynchronously converts a .ppt file to .pptx format with optional configuration.
        /// </summary>
        /// <param name="pptPath">The path to the input .ppt file.</param>
        /// <param name="pptxPath">The path for the output .pptx file.</param>
        /// <param name="options">Optional conversion options and callbacks.</param>
        /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
        /// <returns>A task representing the asynchronous conversion operation.</returns>
        public static async Task ConvertAsync(string pptPath, string pptxPath, ConversionOptions? options, CancellationToken cancellationToken = default)
        {
            // Apply timeout if configured
            if (options?.Timeout > TimeSpan.Zero)
            {
                using var timeoutCts = new CancellationTokenSource(options.Timeout);
                using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);
                
                try
                {
                    await Task.Run(() => Convert(pptPath, pptxPath, options, linkedCts.Token), linkedCts.Token);
                }
                catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
                {
                    throw new TimeoutException($"Conversion timed out after {options.Timeout.TotalSeconds} seconds.");
                }
            }
            else
            {
                await Task.Run(() => Convert(pptPath, pptxPath, options, cancellationToken), cancellationToken);
            }
        }

        /// <summary>
        /// Asynchronously converts a .ppt file to .pptx format and returns detailed result information.
        /// </summary>
        /// <param name="pptPath">The path to the input .ppt file.</param>
        /// <param name="pptxPath">The path for the output .pptx file.</param>
        /// <param name="options">Optional conversion options and callbacks.</param>
        /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
        /// <returns>A task that returns a <see cref="ConversionResult"/> containing conversion details.</returns>
        public static async Task<ConversionResult> ConvertWithResultAsync(string pptPath, string pptxPath, ConversionOptions? options = null, CancellationToken cancellationToken = default)
        {
            var stopwatch = Stopwatch.StartNew();
            var inputFileSize = new FileInfo(pptPath).Length;
            
            try
            {
                await ConvertAsync(pptPath, pptxPath, options, cancellationToken);
                
                stopwatch.Stop();
                var outputFileSize = new FileInfo(pptxPath).Length;
                
                return ConversionResult.SuccessResult(
                    stopwatch.Elapsed,
                    slideCount: 0,
                    imageCount: 0,
                    embeddedResourceCount: 0,
                    inputFileSize,
                    outputFileSize);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                return ConversionResult.FailureResult(ex, stopwatch.Elapsed, inputFileSize: inputFileSize);
            }
        }

        /// <summary>
        /// Converts a .ppt from input stream to .pptx output stream.
        /// </summary>
        /// <param name="inputStream">The input stream containing the .ppt data.</param>
        /// <param name="outputStream">The output stream where .pptx data will be written.</param>
        /// <param name="options">Optional conversion options and callbacks.</param>
        /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
        /// <exception cref="ArgumentNullException">Thrown when input or output stream is null.</exception>
        /// <exception cref="ArgumentException">Thrown when streams are not readable/writable.</exception>
        public static void Convert(Stream inputStream, Stream outputStream, ConversionOptions? options = null, CancellationToken cancellationToken = default)
        {
            if (inputStream == null)
                throw new ArgumentNullException(nameof(inputStream));
            if (outputStream == null)
                throw new ArgumentNullException(nameof(outputStream));
            if (!inputStream.CanRead)
                throw new ArgumentException("Input stream must be readable.", nameof(inputStream));
            if (!outputStream.CanWrite)
                throw new ArgumentException("Output stream must be writable.", nameof(outputStream));

            cancellationToken.ThrowIfCancellationRequested();
            options?.ReportProgress(ConversionPhase.Initializing, 0, "Starting conversion...");

            try
            {
                Presentation presentation;

                // Read phase
                cancellationToken.ThrowIfCancellationRequested();
                options?.ReportProgress(ConversionPhase.Reading, 10, "Reading PPT stream...");
                using (var pptReader = new PptReader(inputStream, options))
                {
                    presentation = pptReader.ReadPresentation(cancellationToken);
                }

                cancellationToken.ThrowIfCancellationRequested();
                options?.ReportProgress(ConversionPhase.ProcessingStructure, 30, $"Found {presentation.Slides.Count} slides...", 0, presentation.Slides.Count);

                // Write phase
                cancellationToken.ThrowIfCancellationRequested();
                options?.ReportProgress(ConversionPhase.Writing, 50, "Writing PPTX stream...");
                using (var pptxWriter = new PptxWriter(outputStream, options))
                {
                    pptxWriter.WritePresentation(presentation, cancellationToken);
                }

                cancellationToken.ThrowIfCancellationRequested();
                options?.ReportProgress(ConversionPhase.Finalizing, 90, "Finalizing...");
                options?.ReportProgress(ConversionPhase.Completed, 100, "Conversion completed successfully.", presentation.Slides.Count, presentation.Slides.Count);
            }
            catch (OperationCanceledException)
            {
                options?.ReportProgress(ConversionPhase.Failed, 0, "Conversion canceled.");
                throw;
            }
            catch (Exception ex)
            {
                options?.ReportProgress(ConversionPhase.Failed, 0, $"Conversion failed: {ex.Message}");
                throw new PptConversionException("Failed to convert stream.", ConversionPhase.Failed, ex);
            }
        }

        /// <summary>
        /// Asynchronously converts a .ppt from input stream to .pptx output stream.
        /// </summary>
        /// <param name="inputStream">The input stream containing the .ppt data.</param>
        /// <param name="outputStream">The output stream where .pptx data will be written.</param>
        /// <param name="options">Optional conversion options and callbacks.</param>
        /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
        /// <returns>A task representing the asynchronous conversion operation.</returns>
        public static Task ConvertAsync(Stream inputStream, Stream outputStream, ConversionOptions? options = null, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => Convert(inputStream, outputStream, options, cancellationToken), cancellationToken);
        }

        private static void ValidatePaths(string pptPath, string pptxPath, ConversionOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(pptPath))
                throw new ArgumentException("Input .ppt path must be provided.", nameof(pptPath));
            if (string.IsNullOrWhiteSpace(pptxPath))
                throw new ArgumentException("Output .pptx/.pptm path must be provided.", nameof(pptxPath));

            if (!File.Exists(pptPath))
                throw new FileNotFoundException("Input .ppt file not found.", pptPath);

            if (Path.GetFullPath(pptPath).Equals(Path.GetFullPath(pptxPath), StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Output path must be different from input path.", nameof(pptxPath));

            // 检查文件大小限制
            if (options?.MaxInputFileSize > 0)
            {
                var fileInfo = new FileInfo(pptPath);
                if (fileInfo.Length > options.MaxInputFileSize)
                {
                    throw new PptConversionException(
                        $"Input file size ({fileInfo.Length} bytes) exceeds maximum allowed size ({options.MaxInputFileSize} bytes).",
                        ConversionPhase.Initializing,
                        pptPath);
                }
            }

            // 检查绝对最大文件大小限制
            var absoluteMaxFileInfo = new FileInfo(pptPath);
            if (absoluteMaxFileInfo.Length > ConversionLimits.AbsoluteMaxInputFileSize)
            {
                throw new PptConversionException(
                    $"Input file size ({absoluteMaxFileInfo.Length} bytes) exceeds absolute maximum allowed size ({ConversionLimits.AbsoluteMaxInputFileSize} bytes).",
                    ConversionPhase.Initializing,
                    pptPath);
            }
        }

        private static void ValidatePresentation(Presentation presentation, ConversionOptions? options)
        {
            if (options?.MaxSlideCount > 0 && presentation.Slides.Count > options.MaxSlideCount)
            {
                throw new PptConversionException(
                    $"Presentation contains {presentation.Slides.Count} slides, which exceeds the maximum allowed ({options.MaxSlideCount}).",
                    ConversionPhase.ProcessingStructure);
            }

            if (options?.MaxImageCount > 0 && presentation.Images.Count > options.MaxImageCount)
            {
                throw new PptConversionException(
                    $"Presentation contains {presentation.Images.Count} images, which exceeds the maximum allowed ({options.MaxImageCount}).",
                    ConversionPhase.ProcessingStructure);
            }
        }
    }
}
