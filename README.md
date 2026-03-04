# NPptToPptx (Nedev.PptToPptx)

NPptToPptx is a fast, lightweight, and standalone .NET library for converting legacy binary PowerPoint presentations (`.ppt`, PPT97-2003 format) into the modern OpenXML format (`.pptx`).

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
*   **Master & Layout Mapping:** Maps elements into standard PresentationML (`.pptx`) relationship structures (e.g., `slideLayout1.xml`).

### 2. Text & Typography
*   **Rich Text Extraction:** Supports extracting text from `ClientTextbox` and matching styling properties from `StyleTextPropAtom`.
*   **Paragraphs & Runs:** Groups text correctly into paragraphs and runs based on original formatting limits.
*   **Encoding Handling:** Translates ANSI and Unicode string formats correctly.

### 3. Images and Graphics
*   **Picture Extraction:** Reads `BStore` and `Blip` data to extract PNG, JPEG, and WMF files directly from the PPT structure.
*   **Shape Grouping:** Identifies basic shape boundaries, positional elements (Top, Left, Width, Height), and `Group` containers.
*   **Escher Properties:** Parses standard `ESCHER_OPT` properties including fill colors and line colors.

### 4. Hyperlinks
*   **Global Link Mapping:** Parses `ExObjList` and `ExHyperlinkAtom` structures to map internal link IDs to target URLs.
*   **Text Run Links:** Automatically maps hyperlink regions to specific text runs and generates standard `rId` relationships in the exported `.pptx`.

### 5. Charts & Data
*   **Embedded OLE Objects:** Detects embedded Microsoft Graph and Excel charting data inside nested OLE storages.
*   **BIFF8 Parsing:** Read underlying `Workbook` and `Book` streams to extract chart data via a lightweight BIFF parser (`PptChartParser`).
*   **Chart Types:** Identifies standard column (`colChart`) and bar variations and parses series, names, category labels, and numeric values.
*   **XML Generation:** Outputs fully valid `chartX.xml` definitions embedded inside the OpenXML package structure.

## Technical Architecture

The codebase is primarily divided into two main processing layers:

1.  **`PptReader.cs`:** The ingestion engine. Uses `OleCompoundFile` to open the binary stream, iterates through the PPT record atoms (e.g., `RT_Slide`, `RT_TextCharsAtom`, `ESCHER_ClientData`), parses the metadata, and hydrates standard intermediate C# domain models (`Models.cs`).
2.  **`PptxWriter.cs`:** The emission engine. Takes the intermediate `Presentation` C# model and writes a valid ZIP-based OpenXML package, generating `[Content_Types].xml`, `_rels`, `slideX.xml`, `chartX.xml`, and dynamically wiring up standard relationships.

## How to Build & Run

Ensure you have the .NET SDK installed (currently targets standard .NET platforms).

```bash
# Clone the repository
git clone <repository_url>
cd NPptToPptx

# Build the project
cd src
dotnet build
```

## Usage

Using the converter in your .NET code is straightforward. Simply call the `Convert` method on the `PptToPptxConverter` class, providing the input `.ppt` file path and the desired output `.pptx` file path.

```csharp
using Nefdev.PptToPptx;

class Program
{
    static void Main(string[] args)
    {
        string inputPpt = @"C:\path\to\legacy\presentation.ppt";
        string outputPptx = @"C:\path\to\output\presentation.pptx";

        // Convert the PPT to PPTX
        PptToPptxConverter.Convert(inputPpt, outputPptx);
        
        System.Console.WriteLine("Conversion complete!");
    }
}
```

## Requirements

*   **.NET SDK:** (e.g., .NET 6, .NET 7, .NET 8, or .NET Standard depending on project target)
*   **No external Office installation required.**

## Known Limitations / Roadmap

While the core elements work, there are areas for future improvement:
*   **Text Styling Enhancements:** Extracting specific font families, exact bold/italic bounds, and intricate text coloring mapping.
*   **Vector Geometry:** Converting complex custom PPT preset shape geometries into standard OpenXML drawing elements.
*   **Tables:** Native extraction of nested table structures and grid spans.
*   **SmartArt:** More complex embedded object and diagram extraction.
*   **Animations and Transitions:** Reading `SlideShow` records to persist slide animations to PPTX.

## Contributing

Contributions are welcome! If you find a bug (e.g., a specific `.ppt` file fails to parse) or want to add support for a missing record type:
1. Fork the repository.
2. Create your feature branch (`git checkout -b feature/MyFeature`).
3. Commit your changes (`git commit -m 'Add some feature'`).
4. Push to the branch (`git push origin feature/MyFeature`).
5. Open a Pull Request.

## License
MIT License (or your chosen associated license)
