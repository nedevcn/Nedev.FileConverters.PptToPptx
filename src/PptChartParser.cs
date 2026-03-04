using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace Nefdev.PptToPptx
{
    public class PptChartParser
    {
        private const ushort BIFF_EOF = 0x000A;
        private const ushort BIFF_BOF = 0x0809;
        // Chart specific records
        private const ushort CH_CHART = 0x1002;
        private const ushort CH_SERIES = 0x1003;
        private const ushort CH_DATAFORMAT = 0x1006;
        private const ushort CH_CHARTFORMAT = 0x1014;
        private const ushort CH_SERIESTEXT = 0x100D;
        
        // Data records
        private const ushort NUMBER = 0x0203;
        private const ushort LABEL = 0x0204;
        private const ushort LABELSST = 0x00FD;
        private const ushort SST = 0x00FC;

        // Common format records
        private const ushort FORMAT = 0x041E;
        
        public Chart ParseChart(byte[] biffData)
        {
            var chart = new Chart();
            chart.Type = "bar"; // Default

            using var stream = new MemoryStream(biffData);
            using var reader = new BinaryReader(stream);

            var sstStrings = new List<string>();
            var sstOffsets = new List<uint>();

            ChartSeries currentSeries = null;
            
            // To align categories and values from Sheet data (often in _123456 Workbook streams)
            // Legacy MS Graph dumps data as simply:
            // Column 0 = Categories
            // Column 1 = Series 1 Values
            // Column 2 = Series 2 Values
            // Row 0 = Series Names
            // Row 1..N = Data values
            
            var cells = new Dictionary<(int row, int col), string>();
            var numbers = new Dictionary<(int row, int col), double>();

            while (stream.Position < stream.Length)
            {
                if (stream.Position + 4 > stream.Length) break;

                ushort recordType = reader.ReadUInt16();
                ushort recordLength = reader.ReadUInt16();

                if (stream.Position + recordLength > stream.Length) break;
                
                long nextPos = stream.Position + recordLength;
                
                try
                {
                    switch (recordType)
                    {
                        case FORMAT:
                            // We could parse formats if needed
                            break;
                            
                        case SST:
                            ParseSstInfo(reader, recordLength, sstStrings, sstOffsets, stream.Position + recordLength);
                            break;
                            
                        case NUMBER:
                            {
                                ushort row = reader.ReadUInt16();
                                ushort col = reader.ReadUInt16();
                                ushort xf = reader.ReadUInt16();
                                double val = reader.ReadDouble();
                                numbers[(row, col)] = val;
                                cells[(row, col)] = val.ToString();
                            }
                            break;
                            
                        case LABEL:
                            {
                                ushort row = reader.ReadUInt16();
                                ushort col = reader.ReadUInt16();
                                ushort xf = reader.ReadUInt16();
                                ushort strLen = reader.ReadUInt16();
                                if (strLen > 0)
                                {
                                    // BIFF8 string format: 1 byte flags (bit 0 = 1 for unicode)
                                    byte flags = reader.ReadByte();
                                    bool isUnicode = (flags & 0x01) == 1;
                                    string text;
                                    
                                    if (isUnicode)
                                        text = Encoding.Unicode.GetString(reader.ReadBytes(strLen * 2));
                                    else
                                        text = Encoding.GetEncoding(1252).GetString(reader.ReadBytes(strLen));
                                        
                                    cells[(row, col)] = text;
                                }
                            }
                            break;
                            
                        case LABELSST:
                            {
                                ushort row = reader.ReadUInt16();
                                ushort col = reader.ReadUInt16();
                                ushort xf = reader.ReadUInt16();
                                uint sstIndex = reader.ReadUInt32();
                                if (sstIndex < sstStrings.Count)
                                {
                                    cells[(row, col)] = sstStrings[(int)sstIndex];
                                }
                            }
                            break;
                            
                        case CH_CHARTFORMAT:
                            // Not fully detailed, but can extract type hints or rely on default
                            break;
                    }
                }
                catch
                {
                    // Ignore malformed records
                }

                stream.Position = nextPos;
            }

            // After parsing all cells, assemble the Chart
            // Row 0, Col 1..MaxCol -> Series Names
            // Row 1..MaxRow, Col 0 -> Categories
            // Row 1..MaxRow, Col 1..MaxCol -> Values
            
            int maxRow = -1;
            int maxCol = -1;
            foreach (var key in cells.Keys)
            {
                maxRow = Math.Max(maxRow, key.row);
                maxCol = Math.Max(maxCol, key.col);
            }

            if (maxRow >= 0 && maxCol >= 0)
            {
                // Find categories
                var categoryList = new List<string>();
                for (int r = 1; r <= maxRow; r++)
                {
                    if (cells.TryGetValue((r, 0), out string cat))
                        categoryList.Add(cat);
                    else
                        categoryList.Add("");
                }

                // Build series
                for (int c = 1; c <= maxCol; c++)
                {
                    var series = new ChartSeries();
                    
                    if (cells.TryGetValue((0, c), out string sName))
                        series.Name = sName;
                    else
                        series.Name = $"Series {c}";
                        
                    series.Categories = new List<string>(categoryList);
                    
                    for (int r = 1; r <= maxRow; r++)
                    {
                        if (numbers.TryGetValue((r, c), out double num))
                            series.Values.Add(num);
                        else
                            series.Values.Add(0.0);
                    }
                    
                    if (series.Values.Count > 0)
                    {
                        chart.Series.Add(series);
                    }
                }
            }

            // Fallback if empty (shouldn't happen for valid MS Graph)
            if (chart.Series.Count == 0)
            {
                var dummySeries = new ChartSeries { Name = "Series 1" };
                dummySeries.Categories.Add("Category 1");
                dummySeries.Values.Add(1.0);
                chart.Series.Add(dummySeries);
            }

            return chart;
        }

        private void ParseSstInfo(BinaryReader reader, ushort length, List<string> strings, List<uint> offsets, long endPosition)
        {
            if (length < 8) return;
            
            uint totalStrings = reader.ReadUInt32();
            uint uniqueStrings = reader.ReadUInt32();
            
            for (int i = 0; i < uniqueStrings && reader.BaseStream.Position < endPosition; i++)
            {
                try
                {
                    ushort charCount = reader.ReadUInt16();
                    byte flags = reader.ReadByte();
                    
                    bool isUnicode = (flags & 0x01) == 1;
                    bool hasExtString = (flags & 0x04) == 4;
                    bool hasRichText = (flags & 0x08) == 8;
                    
                    ushort runCount = 0;
                    if (hasRichText) runCount = reader.ReadUInt16();
                    
                    uint extLength = 0;
                    if (hasExtString) extLength = reader.ReadUInt32();
                    
                    string text = "";
                    if (charCount > 0)
                    {
                        if (isUnicode)
                            text = Encoding.Unicode.GetString(reader.ReadBytes(charCount * 2));
                        else
                            text = Encoding.GetEncoding(1252).GetString(reader.ReadBytes(charCount));
                    }
                    
                    strings.Add(text);
                    
                    // Skip formatting runs and extended info
                    if (hasRichText) reader.BaseStream.Position += runCount * 4;
                    if (hasExtString) reader.BaseStream.Position += extLength;
                }
                catch
                {
                    // If parsing a string fails, try to salvage
                    if (strings.Count == i) strings.Add("");
                }
            }
        }
    }
}
