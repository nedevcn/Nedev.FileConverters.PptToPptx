# Nedev.FileConverters.PptToPptx

A library and command-line tool for converting legacy PowerPoint `.ppt` files to modern OpenXML `.pptx` (or `.pptm`) format.

NPptToPptx is a fast, lightweight .NET **library** for converting legacy binary PowerPoint presentations (`.ppt`, PPT97-2003 format) into the modern OpenXML format (`.pptx`).

The repository contains two separate projects:

* `Nedev.FileConverters.PptToPptx` – the core library (multi-targeted).
* `Nedev.FileConverters.PptToPptx.Cli` – a small console application that wraps the library and is published as a global tool.

Unlike many other solutions, this project works by directly parsing the underlying OLE Compound File streams and binary records (such as BIFF records for charts and ESCHER records for drawings), eliminating the need for Office Interop or heavy third-party dependencies.

## Key Advantages
*   **No Dependency on Microsoft Office:** Runs anywhere .NET runs (Windows, Linux, macOS).
*   **Native Parsing:** Reads the proprietary `.ppt` binary format at the byte level.
*   **Lightweight:** Minimal memory footprint compared to headless Office instances or massive third-party commercial SDKs.

## Features Currently Implemented

The parser and writer have been significantly developed to handle a wide range of standard presentation elements:

### 1. Structure & Layout Definition
*   **OLE Compound File Parsing:** Custom implementation to extract the `PowerPoint Document`, `Pictures`, and nested OLE storages (`_1326458456` etc.).
*   **Slide & Shape Containers:** Correctly parses sequence of slides, sizes, drawing groups, and slide containers.
*   **Master & Layout Mapping:** Writes a minimal but standard PresentationML structure (slide masters, layouts, theme, rels). Layouts include basic placeholders (title/body) to keep the package closer to PowerPoint defaults.
*   **Slide Notes:** Writes `notesSlideX.xml` parts when notes are present.

### 2. Text & Typography
*   **Rich Text Extraction:** Supports extracting text from `ClientTextbox` and matching styling properties from `StyleTextPropAtom`.
*   **Paragraphs & Runs:** Groups text correctly into paragraphs and runs based on original formatting limits.
*   **Bullets & Indentation (best-effort):** Maps bullet character and outline indentation level when present.
*   **Paragraph Spacing (best-effort):** Attempts to persist line spacing and space before/after.
*   **Encoding Handling:** Translates ANSI and Unicode string formats correctly.

### 3. Images and Graphics
*   **Picture Extraction:** Reads `BStore` and `Blip` data to extract PNG, JPEG, EMF/WMF, BMP, TIFF (and others) directly from the PPT structure.
*   **Shape Grouping:** Identifies basic shape boundaries, positional elements (Top, Left, Width, Height), and `Group` containers.
*   **Escher Properties:** Parses common `ESCHER_OPT` properties including fill colors and line colors.
*   **Line & Fill Effects (best-effort):** Supports line width, simple dash styles, a basic 2-stop gradient fill, and a simplified shadow mapping.

### 4. Hyperlinks
*   **Global Link Mapping:** Parses `ExObjList` and `ExHyperlinkAtom` structures to map internal link IDs to target URLs.
*   **Text Run Links:** Automatically maps hyperlink regions to specific text runs and generates standard `rId` relationships in the exported `.pptx`.
*   **Internal Jump Actions (best-effort):** Parses `InteractiveInfoAtom` to preserve common “jump” actions (next/prev/first/last slide, end show) via `ppaction://...` actions where possible.

### 5. Charts & Data
*   **Embedded OLE Objects:** Detects embedded Microsoft Graph and Excel charting data inside nested OLE storages.
*   **BIFF8 Parsing:** Read underlying `Workbook` and `Book` streams to extract chart data via a lightweight BIFF parser (`PptChartParser`).
*   **Chart Types:** Identifies common chart types (bar/line/pie/area/scatter/radar) and parses series, names, category labels, numeric values, and some formatting hints.
*   **XML Generation:** Outputs fully valid `chartX.xml` definitions embedded inside the OpenXML package structure.

### 6. Tables (best-effort)
*   **Table Detection:** Detects native tables via Programmable Tags (`___PPT10` / `___PPT12`) and a simple grid heuristic.
*   **Table Output:** Writes tables as DrawingML `<a:tbl>` with cell text, simple fill colors, and basic margins/vertical alignment.

### 7. Embedded Objects / Media (minimal preservation)
*   **OLE Payload Extraction (best-effort):** For non-chart embedded objects, attempts to extract a payload stream from the nested OLE storage.
*   **Package Embeddings:** Stores extracted payloads under `ppt/embeddings/oleObjectN.bin` so information isn’t silently lost (PowerPoint may or may not render these as playable media without richer OOXML structures).

### 8. Transitions & Animations (best-effort)
*   **Slide Transitions:** Maps a subset of `SlideShowSlideInfoAtom` transitions to `<p:transition>`.
*   **Shape Animations:** Writes a simplified `<p:timing>` tree for basic entrance effects when animation metadata is present.

### 9. VBA / Macros (packaging support)
*   **VBA Stream Extraction:** Reads `VBA Project` from OLE and writes `ppt/vba/vbaProject.bin`.
*   **Macro-Enabled Packaging:** Adds the required relationship and content-type entries so macro content can be recognized.  
    - Recommended output extension for macro content: **`.pptm`**.

## Technical Architecture

The codebase is primarily divided into two main processing layers:

1.  **`PptReader.cs`:** The ingestion engine. Uses `OleCompoundFile` to open the binary stream, iterates through the PPT record atoms (e.g., `RT_Slide`, `RT_TextCharsAtom`, `ESCHER_ClientData`), parses the metadata, and hydrates standard intermediate C# domain models (`Models.cs`).
2.  **`PptxWriter.cs`:** The emission engine. Takes the intermediate `Presentation` C# model and writes a valid ZIP-based OpenXML package, generating `[Content_Types].xml`, `_rels`, `slideX.xml`, `chartX.xml`, and dynamically wiring up standard relationships.

## How to Build & Run

Ensure you have the .NET SDK (8.0+) installed.

```bash
# Clone the repository
git clone <repository_url>
cd Nedev.FileConverters.PptToPptx

# Build the library (multi-targets)
dotnet build src/Nedev.FileConverters.PptToPptx

# Build the CLI tool separately
dotnet build src/Nedev.FileConverters.PptToPptx.Cli
```

After building the CLI project you can run the executable under `src/Nedev.FileConverters.PptToPptx.Cli/bin/Debug/net8.0` or install it as a tool.

## Usage

Using the converter in your .NET code is straightforward. Simply call the `Convert` method on the `PptToPptxConverter` class, providing the input `.ppt` file path and the desired output `.pptx` file path.

```csharp
using Nedev.FileConverters.PptToPptx;
using Nedev.FileConverters.Core;

class Program
{
    static void Main(string[] args)
    {
        string inputPpt = @"C:\path\to\legacy\presentation.ppt";
        string outputPptx = @"C:\path\to\output\presentation.pptx";

        // configure options from the shared core library
        var opts = new ConversionOptions { KeepTempFiles = true };

        // Convert the PPT to PPTX
        PptToPptxConverter.Convert(inputPpt, outputPptx, opts);
        
        System.Console.WriteLine("Conversion complete!");
    }
}
```

### Macro-enabled output (`.pptm`)

If the source `.ppt` contains a VBA project, you should write to a `.pptm` path so PowerPoint treats the output as macro-enabled:

```csharp
PptToPptxConverter.Convert(@"C:\in\legacy.ppt", @"C:\out\converted.pptm");
```

## Command-line tool

A standalone CLI project lives in `src/Nedev.FileConverters.PptToPptx.Cli`.  Build and run it separately (this project also references the core package for the registry API):

```bash
cd src/Nedev.FileConverters.PptToPptx.Cli
dotnet run -- in.ppt out.pptx
```

Or install the tool after packing:

```bash
dotnet tool install --global Nedev.FileConverters.PptToPptx.Cli --version 1.0.0
ppt2pptx in.ppt out.pptx
```

## Requirements

*   **.NET SDK:** .NET 8 (current project target)
*   **No external Office installation required.**

## Dependencies

This converter depends on a shared core library available as the NuGet package **Nedev.FileConverters.Core** (version `0.1.0` or later).  The package provides a lightweight framework for registering and discovering file‑format converters.  Our library integrates by providing an `IFileConverter` adapter annotated with `[FileConverter("ppt","pptx")]`, so consumers who reference the core package can use the generic `Nedev.FileConverters.Converter` API or the DI helpers without additional code.  (Conversion options and encoding helpers remain defined locally to avoid a hard dependency on new package versions.)

## NuGet Package

The library project produces a NuGet package when built (`<GeneratePackageOnBuild>` is enabled).
The package ID is **Nedev.FileConverters.PptToPptx** and includes the README as the package readme; the licence is MIT.

To create the library package manually:

```bash
cd src/Nedev.FileConverters.PptToPptx
dotnet pack -c Release
```

Install it locally or publish to NuGet.org as usual.

The CLI project also packs itself and is intended to be deployed as a .NET global tool. Build or pack it independently:

```bash
cd src/Nedev.FileConverters.PptToPptx.Cli
dotnet pack -c Release    # produces Nedev.FileConverters.PptToPptx.Cli.1.0.0.nupkg
```

Then install the tool:

```bash
dotnet tool install --global Nedev.FileConverters.PptToPptx.Cli --version 1.0.0
ppt2pptx in.ppt out.pptx
```


## Known Limitations / Roadmap

While the core elements work, there are areas for future improvement:
*   **Text Fidelity:** Bullet numbering styles (`<a:buAutoNum>`), precise unit conversion for spacing/indent, and full paragraph/run formatting coverage.
*   **Layouts & Placeholders:** Full placeholder semantics (date/footer/slide number/header) and multiple layout types beyond title/body.
*   **Embedded Media:** Proper OOXML audio/video parts and relationships (currently payloads are stored in `ppt/embeddings` for minimal preservation).
*   **OLE Objects:** Full round-trip of embedded objects beyond charts (icon/preview rendering, activation verbs, etc.).
*   **Vector/Shape Effects:** More complete mapping for gradients, patterns, transparency, arrows, 3D, glow, etc.
*   **Animations/Interactions:** More accurate timing trees, triggers, and action buttons (current output is best-effort).
*   **SmartArt / Diagrams:** Not implemented.

## Contributing

Contributions are welcome! If you find a bug (e.g., a specific `.ppt` file fails to parse) or want to add support for a missing record type:
1. Fork the repository.
2. Create your feature branch (`git checkout -b feature/MyFeature`).
3. Commit your changes (`git commit -m 'Add some feature'`).
4. Push to the branch (`git push origin feature/MyFeature`).
5. Open a Pull Request.

## License
This project is licensed under the [MIT License](LICENSE).
