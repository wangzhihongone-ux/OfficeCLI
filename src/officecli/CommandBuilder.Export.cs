// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json.Nodes;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildExportCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Office document path (.docx)" };
        var formatOpt = new Option<string>("--format")
        {
            Description = "Export format: markdown, md, json",
            DefaultValueFactory = _ => "markdown"
        };
        var outOpt = new Option<string?>("--out", "-o") { Description = "Write output to a file instead of stdout" };

        var command = new Command("export", "Export an Office document to Markdown or structured JSON");
        command.Add(fileArg);
        command.Add(formatOpt);
        command.Add(outOpt);
        command.Add(jsonOption);

        command.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var format = NormalizeExportFormat(result.GetValue(formatOpt));
            var outPath = result.GetValue(outOpt);

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext != ".docx")
                throw new CliException($"export currently supports .docx only (got {ext})")
                {
                    Code = "unsupported_format",
                    Suggestion = "Use a .docx file. Excel export will be added as a separate pass.",
                    ValidValues = [".docx"]
                };

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            if (handler is not WordHandler word)
                throw new CliException("export currently supports Word documents only.")
                {
                    Code = "unsupported_type",
                    ValidValues = [".docx"]
                };

            var output = format switch
            {
                "markdown" => word.ExportAsMarkdown(),
                "json" => word.ExportAsJson(),
                _ => throw new CliException($"Unsupported --format: {format}. Valid: markdown, json")
                {
                    Code = "invalid_format",
                    ValidValues = ["markdown", "md", "json"]
                }
            };

            if (outPath == "-") outPath = null;
            if (!string.IsNullOrEmpty(outPath))
            {
                File.WriteAllText(outPath, output);
                if (json)
                {
                    var meta = new JsonObject
                    {
                        ["outputFile"] = outPath,
                        ["format"] = format
                    };
                    Console.WriteLine(OutputFormatter.WrapEnvelope(meta.ToJsonString()));
                }
                else
                    Console.WriteLine(Path.GetFullPath(outPath));
            }
            else
            {
                if (json && format != "json")
                    Console.WriteLine(OutputFormatter.WrapEnvelopeText(output));
                else if (json)
                    Console.WriteLine(OutputFormatter.WrapEnvelope(output));
                else
                    Console.Write(output);
            }

            return 0;
        }, json); });

        return command;
    }

    private static string NormalizeExportFormat(string? format)
    {
        var f = (format ?? "markdown").Trim().ToLowerInvariant();
        return f switch
        {
            "" or "markdown" or "md" => "markdown",
            "json" => "json",
            _ => f
        };
    }
}
