# DOCX Export Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a full-fidelity `officecli export` command that converts `.docx` files to Markdown or structured JSON.

**Architecture:** Keep the command layer thin in `CommandBuilder.Export.cs` and put Word-specific conversion in a new `WordHandler.Export.cs` partial. Markdown output uses embedded HTML tables so nested Word tables can be represented without losing structure.

**Tech Stack:** C# / .NET 10, System.CommandLine, DocumentFormat.OpenXml, existing OfficeCLI command and Word handler patterns.

---

### Task 1: Register Export Command

**Files:**
- Create: `src/officecli/CommandBuilder.Export.cs`
- Modify: `src/officecli/CommandBuilder.cs`

- [x] Add `BuildExportCommand(jsonOption)` and register it in the root command.
- [x] Accept `file`, `--format markdown|md|json`, and `--out|-o`.
- [x] Route `.docx` files to `WordHandler.ExportAsMarkdown()` or `WordHandler.ExportAsJson()`.

### Task 2: Implement Word Exporter

**Files:**
- Create: `src/officecli/Handlers/Word/WordHandler.Export.cs`

- [x] Walk body elements in document order.
- [x] Convert heading paragraphs to Markdown headings.
- [x] Convert ordinary paragraphs to Markdown paragraphs.
- [x] Convert Word tables to HTML tables in Markdown.
- [x] Preserve nested table structure recursively in JSON and Markdown.
- [x] Emit structured JSON with `metadata`, `outline`, and `blocks`.

### Task 3: Verify Locally

**Files:**
- Use sample `.docx` files from `examples/word` and `assets/showcase`.

- [x] Run static text checks with `rg`.
- [x] Run `dotnet build src/officecli/officecli.csproj`.
- [x] Run `officecli export <sample.docx> --format markdown`.
- [x] Run `officecli export <sample.docx> --format json`.

Note: `officecli.slnx` references `tests/OfficeCli.Tests/OfficeCli.Tests.csproj`, which is not present in this checkout. The main CLI project builds cleanly.
