// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Net;
using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    public string ExportAsMarkdown()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "";

        var blocks = BuildExportBlocks(body);
        var sb = new StringBuilder();
        foreach (var block in blocks)
        {
            var markdown = BlockToMarkdown(block);
            if (string.IsNullOrWhiteSpace(markdown)) continue;
            if (sb.Length > 0) sb.AppendLine().AppendLine();
            sb.Append(markdown.TrimEnd());
        }
        if (sb.Length > 0) sb.AppendLine();
        return sb.ToString();
    }

    public string ExportAsJson()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null)
        {
            return new JsonObject
            {
                ["metadata"] = new JsonObject
                {
                    ["fileName"] = Path.GetFileName(_filePath),
                    ["paragraphs"] = 0,
                    ["tables"] = 0,
                    ["headings"] = 0
                },
                ["outline"] = new JsonArray(),
                ["blocks"] = new JsonArray()
            }.ToJsonString(OutputFormatter.PublicJsonOptions);
        }

        var blocks = BuildExportBlocks(body);
        var outline = new JsonArray();
        foreach (var block in FlattenBlocks(blocks))
        {
            if (block.Type != "heading") continue;
            outline.Add((JsonNode)new JsonObject
            {
                ["path"] = block.Path,
                ["level"] = block.Level,
                ["text"] = block.Text,
                ["style"] = block.Style
            });
        }

        var root = new JsonObject
        {
            ["metadata"] = new JsonObject
            {
                ["fileName"] = Path.GetFileName(_filePath),
                ["paragraphs"] = FlattenBlocks(blocks).Count(b => b.Type is "paragraph" or "heading"),
                ["tables"] = FlattenBlocks(blocks).Count(b => b.Type == "table"),
                ["headings"] = FlattenBlocks(blocks).Count(b => b.Type == "heading")
            },
            ["outline"] = outline,
            ["blocks"] = BlocksToJson(blocks)
        };
        return root.ToJsonString(OutputFormatter.PublicJsonOptions);
    }

    private List<WordExportBlock> BuildExportBlocks(Body body)
    {
        return BuildExportBlocks(GetBodyElements(body), "/body");
    }

    private List<WordExportBlock> BuildExportBlocks(IEnumerable<OpenXmlElement> elements, string pathPrefix)
    {
        var blocks = new List<WordExportBlock>();
        int pIdx = 0, tblIdx = 0, eqIdx = 0;

        foreach (var element in elements)
        {
            if (element is Paragraph para)
            {
                pIdx++;
                blocks.Add(BuildParagraphExportBlock(para, $"{pathPrefix}/{BuildParaPathSegment(para, pIdx)}"));
            }
            else if (element is Table table)
            {
                tblIdx++;
                blocks.Add(BuildTableExportBlock(table, $"{pathPrefix}/tbl[{tblIdx}]"));
            }
            else if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                eqIdx++;
                blocks.Add(new WordExportBlock
                {
                    Type = "equation",
                    Path = $"{pathPrefix}/oMathPara[{eqIdx}]",
                    Text = FormulaParser.ToReadableText(element)
                });
            }
        }

        return blocks;
    }

    private WordExportBlock BuildParagraphExportBlock(Paragraph para, string path)
    {
        var oMathParaChild = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
        if (oMathParaChild != null)
        {
            return new WordExportBlock
            {
                Type = "equation",
                Path = path,
                Text = FormulaParser.ToReadableText(oMathParaChild)
            };
        }

        var styleName = GetStyleName(para);
        var headingLevel = IsHeadingStyleForExport(styleName) ? GetHeadingLevel(styleName) : (int?)null;
        var text = BuildParagraphExportText(para);

        return new WordExportBlock
        {
            Type = headingLevel.HasValue ? "heading" : "paragraph",
            Path = path,
            Text = text,
            Style = styleName,
            Level = headingLevel
        };
    }

    private WordExportBlock BuildTableExportBlock(Table table, string path)
    {
        var block = new WordExportBlock
        {
            Type = "table",
            Path = path,
            Rows = new List<WordExportRow>()
        };

        int rowIdx = 0;
        foreach (var row in table.Elements<TableRow>())
        {
            rowIdx++;
            var exportRow = new WordExportRow { Cells = new List<WordExportCell>() };
            int cellIdx = 0;
            foreach (var cell in row.Elements<TableCell>())
            {
                cellIdx++;
                var cellPath = $"{path}/tr[{rowIdx}]/tc[{cellIdx}]";
                var cellBlocks = BuildExportBlocks(cell.ChildElements, cellPath);
                var props = cell.TableCellProperties;
                var gridSpan = props?.GridSpan?.Val?.Value;
                int? colSpan = null;
                if (gridSpan != null && int.TryParse(gridSpan.ToString(), out var span))
                    colSpan = span;
                var vMerge = props?.VerticalMerge == null
                    ? null
                    : props.VerticalMerge.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
                exportRow.Cells.Add(new WordExportCell
                {
                    Path = cellPath,
                    Text = BuildCellText(cellBlocks),
                    Blocks = cellBlocks,
                    ColSpan = colSpan,
                    VerticalMerge = vMerge
                });
            }
            block.Rows.Add(exportRow);
        }

        return block;
    }

    private string BuildParagraphExportText(Paragraph para)
    {
        var fieldSentinelText = TryGetParagraphTextWithFieldSentinels(para);
        if (fieldSentinelText != null)
            return GetListPrefix(para) + fieldSentinelText;

        var ffText = GetParagraphTextWithFormFields(para);
        var mathElements = FindMathElements(para);
        string text;
        if (ffText != null)
            text = ffText;
        else if (mathElements.Count > 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
            text = string.Concat(mathElements.Select(FormulaParser.ToReadableText));
        else if (mathElements.Count > 0)
            text = GetParagraphTextWithMath(para);
        else
            text = GetParagraphText(para);

        var extras = new List<string>();
        foreach (var drawing in para.Descendants<Drawing>())
            extras.Add($"[Image: {GetDrawingInfo(drawing)}]");
        foreach (var embObj in para.Descendants<EmbeddedObject>())
        {
            var oleEl = embObj.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
            var progId = oleEl?.GetAttributes().FirstOrDefault(a => a.LocalName == "ProgID").Value;
            extras.Add($"[OLE: {(string.IsNullOrEmpty(progId) ? "Object" : progId)}]");
        }

        if (extras.Count > 0)
            text = string.IsNullOrWhiteSpace(text) ? string.Join(" ", extras) : $"{text} {string.Join(" ", extras)}";

        return GetListPrefix(para) + text;
    }

    private static bool IsHeadingStyleForExport(string styleName)
    {
        return styleName.Contains("Heading", StringComparison.OrdinalIgnoreCase)
            || styleName.Contains("标题", StringComparison.OrdinalIgnoreCase)
            || styleName.Equals("Title", StringComparison.OrdinalIgnoreCase)
            || styleName.Equals("Subtitle", StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildCellText(List<WordExportBlock> blocks)
    {
        return string.Join("\n", FlattenBlocks(blocks)
            .Where(b => b.Type != "table")
            .Select(b => b.Text)
            .Where(t => !string.IsNullOrWhiteSpace(t)));
    }

    private static IEnumerable<WordExportBlock> FlattenBlocks(IEnumerable<WordExportBlock> blocks)
    {
        foreach (var block in blocks)
        {
            yield return block;
            if (block.Rows == null) continue;
            foreach (var cell in block.Rows.SelectMany(r => r.Cells))
            foreach (var child in FlattenBlocks(cell.Blocks))
                yield return child;
        }
    }

    private static JsonArray BlocksToJson(IEnumerable<WordExportBlock> blocks)
    {
        var array = new JsonArray();
        foreach (var block in blocks)
            array.Add((JsonNode)BlockToJson(block));
        return array;
    }

    private static JsonObject BlockToJson(WordExportBlock block)
    {
        var obj = new JsonObject
        {
            ["type"] = block.Type,
            ["path"] = block.Path
        };
        if (!string.IsNullOrEmpty(block.Text)) obj["text"] = block.Text;
        if (!string.IsNullOrEmpty(block.Style)) obj["style"] = block.Style;
        if (block.Level.HasValue) obj["level"] = block.Level.Value;
        if (block.Rows != null)
        {
            var rows = new JsonArray();
            foreach (var row in block.Rows)
            {
                var cells = new JsonArray();
                foreach (var cell in row.Cells)
                {
                    var cellObj = new JsonObject
                    {
                        ["path"] = cell.Path,
                        ["text"] = cell.Text,
                        ["blocks"] = BlocksToJson(cell.Blocks)
                    };
                    if (cell.ColSpan.HasValue) cellObj["colSpan"] = cell.ColSpan.Value;
                    if (!string.IsNullOrEmpty(cell.VerticalMerge)) cellObj["verticalMerge"] = cell.VerticalMerge;
                    cells.Add((JsonNode)cellObj);
                }
                rows.Add((JsonNode)new JsonObject { ["cells"] = cells });
            }
            obj["rows"] = rows;
        }
        return obj;
    }

    private static string BlockToMarkdown(WordExportBlock block)
    {
        return block.Type switch
        {
            "heading" => $"{new string('#', Math.Clamp(block.Level ?? 1, 1, 6))} {NormalizeMarkdownText(block.Text)}",
            "paragraph" => NormalizeMarkdownText(block.Text),
            "equation" => $"`{NormalizeMarkdownText(block.Text)}`",
            "table" => TableToHtml(block),
            _ => NormalizeMarkdownText(block.Text)
        };
    }

    private static string TableToHtml(WordExportBlock table)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<table>");
        foreach (var row in table.Rows ?? Enumerable.Empty<WordExportRow>())
        {
            sb.AppendLine("  <tr>");
            foreach (var cell in row.Cells)
            {
                var attrs = new List<string>();
                if (cell.ColSpan.HasValue && cell.ColSpan.Value > 1) attrs.Add($"colspan=\"{cell.ColSpan.Value}\"");
                if (!string.IsNullOrEmpty(cell.VerticalMerge)) attrs.Add($"data-vmerge=\"{Html(cell.VerticalMerge)}\"");
                var attrText = attrs.Count > 0 ? " " + string.Join(" ", attrs) : "";
                sb.Append($"    <td{attrText}>");
                var content = CellBlocksToHtml(cell.Blocks);
                if (content.Contains('\n'))
                    sb.AppendLine().Append(content).AppendLine("    </td>");
                else
                    sb.Append(content).AppendLine("</td>");
            }
            sb.AppendLine("  </tr>");
        }
        sb.Append("</table>");
        return sb.ToString();
    }

    private static string CellBlocksToHtml(List<WordExportBlock> blocks)
    {
        var parts = new List<string>();
        foreach (var block in blocks)
        {
            if (block.Type == "table")
            {
                parts.Add(TableToHtml(block));
                continue;
            }

            var text = NormalizeMarkdownText(block.Text);
            if (string.IsNullOrWhiteSpace(text)) continue;
            if (block.Type == "heading")
                parts.Add($"<p><strong>{Html(text)}</strong></p>");
            else if (block.Type == "equation")
                parts.Add($"<p><code>{Html(text)}</code></p>");
            else
                parts.Add($"<p>{Html(text)}</p>");
        }
        return string.Join("\n", parts);
    }

    private static string NormalizeMarkdownText(string? text)
    {
        return (text ?? "").Replace("\r\n", "\n").Replace('\r', '\n').Trim();
    }

    private static string Html(string? text) => WebUtility.HtmlEncode(text ?? "");

    private sealed class WordExportBlock
    {
        public string Type { get; set; } = "";
        public string Path { get; set; } = "";
        public string? Text { get; set; }
        public string? Style { get; set; }
        public int? Level { get; set; }
        public List<WordExportRow>? Rows { get; set; }
    }

    private sealed class WordExportRow
    {
        public List<WordExportCell> Cells { get; set; } = new();
    }

    private sealed class WordExportCell
    {
        public string Path { get; set; } = "";
        public string Text { get; set; } = "";
        public List<WordExportBlock> Blocks { get; set; } = new();
        public int? ColSpan { get; set; }
        public string? VerticalMerge { get; set; }
    }
}
