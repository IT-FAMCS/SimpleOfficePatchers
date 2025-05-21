#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.JavaScript;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleOfficePatchers.Models;

namespace SimpleOfficePatchers.Patchers;

public partial class WordPatcher
{
    [JSExport]
    public static byte[] PatchDocument(byte[] documentBytes, string serializedPatches)
    {
        var patches =
            JsonSerializer.Deserialize(serializedPatches, WordPatchesContext.Default.WordPatches);
        ArgumentNullException.ThrowIfNull(patches);
        var stream = new MemoryStream();
        stream.Write(documentBytes);
        var document = WordprocessingDocument.Open(stream, true);

        foreach (var (placeholder, patch) in patches.Text)
            ApplyTextPatch(document,
                FindParagraphsWithPlaceholder(document, placeholder), patch);
        foreach (var (placeholder, patch) in patches.List)
            ApplyListPatch(document, FindParagraphsWithPlaceholder(document, placeholder), patch);

        document.Dispose();
        return stream.ToArray();
    }

    [JSExport]
    public static string ExtractPlaceholders(byte[] documentBytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        var body = document.MainDocumentPart?.Document.Body;
        ArgumentNullException.ThrowIfNull(body);

        var textMatches = TextPlaceholderRegex().Matches(body.InnerText);
        var listMatches = ListPlaceholderRegex().Matches(body.InnerText);

        var duplicatePlaceholders = textMatches
            .Where(tm => listMatches.Any(lm => lm.Groups[1].Value == tm.Groups[1].Value))
            .Select(m => m.Groups[1].Value).ToList();
        if (duplicatePlaceholders.Count != 0)
            throw new Exception(
                $"duplicate (both text and list types) placeholders found: {string.Join(", ", duplicatePlaceholders)}");

        var placeholders = (from textMatch in textMatches
                let name = textMatch.Groups[1].Value
                let description = textMatch.Groups[2].Value
                select new WordPlaceholder(textMatch.Value, name, description.Trim(), false))
            .DistinctBy(m => m.Raw)
            .ToList();
        placeholders.AddRange((from listMatch in listMatches
                let name = listMatch.Groups[1].Value
                let description = listMatch.Groups[2].Value
                select new WordPlaceholder(listMatch.Value, name, description.Trim(), true))
            .DistinctBy(m => m.Raw));

        return JsonSerializer.Serialize(placeholders,
            WordPlaceholderContext.Default.ListWordPlaceholder);
    }

    private record PlaceholderInformation((int, int) RunIndices, Paragraph SourceParagraph);

    private static void ApplyListPatch(WordprocessingDocument document,
        List<PlaceholderInformation>? placeholdersInformation,
        WordListPatch patch)
    {
        if (patch.Patches.Count == 0) return;
        var body = document.MainDocumentPart?.Document.Body;
        ArgumentNullException.ThrowIfNull(placeholdersInformation);
        ArgumentNullException.ThrowIfNull(body);

        foreach (var placeholderInformation in placeholdersInformation)
        {
            var updatedFirstParagraph = ReplaceParagraphText(placeholderInformation, patch.Patches.First().Text);
            body.ReplaceChild(updatedFirstParagraph, placeholderInformation.SourceParagraph);

            foreach (var text in patch.Patches.Skip(1).Reverse().Select(p => p.Text))
            {
                var newParagraph = ReplaceParagraphText(placeholderInformation, text);
                body.InsertAfter(newParagraph, updatedFirstParagraph);
            }
        }
    }

    private static void ApplyTextPatch(WordprocessingDocument document,
        List<PlaceholderInformation>? placeholdersInformation,
        WordTextPatch patch)
    {
        var body = document.MainDocumentPart?.Document.Body;
        ArgumentNullException.ThrowIfNull(placeholdersInformation);
        ArgumentNullException.ThrowIfNull(body);

        foreach (var placeholderInformation in placeholdersInformation)
        {
            var newParagraph = ReplaceParagraphText(placeholderInformation, patch.Text);
            body.ReplaceChild(newParagraph, placeholderInformation.SourceParagraph);
        }
    }

    private static List<PlaceholderInformation> FindParagraphsWithPlaceholder(WordprocessingDocument document,
        string placeholder)
    {
        var body = document.MainDocumentPart?.Document.Body;
        ArgumentNullException.ThrowIfNull(body);

        var result = new List<PlaceholderInformation>();
        foreach (var child in body.ChildElements)
        {
            if (child is Paragraph paragraph && paragraph.InnerText.Contains(placeholder))
                result.Add(new PlaceholderInformation(FindRunIndices(paragraph, placeholder), paragraph));
        }

        return result;
    }

    private static Paragraph ReplaceParagraphText(PlaceholderInformation placeholderInformation, string with)
    {
        var paragraph = (Paragraph)placeholderInformation.SourceParagraph.Clone();
        var (startIndex, endIndex) = placeholderInformation.RunIndices;

        var run = (Run)paragraph.ChildElements[startIndex];
        foreach (var descendant in run.Descendants<Text>()) descendant.Remove();
        run.Append(new Text(with));

        var removedRuns = Enumerable.Range(startIndex + 1, endIndex - startIndex)
            .Select(i =>
                (Run)paragraph.ChildElements[i]);
        // i have zero idea why but items have to be removed in reverse for this to work
        foreach (var removedRun in removedRuns.Reverse()) removedRun.Remove();

        return paragraph;
    }

    private static (int, int) FindRunIndices(Paragraph paragraph, string placeholder)
    {
        int currentStart = -1, currentEnd = -1, currentStop = 0;
        foreach (var (index, child) in paragraph.ChildElements.Index())
        {
            if (child is not Run run) continue;
            var text = run.GetFirstChild<Text>()?.InnerText;
            if (text is null) continue;
            if (text.Contains(placeholder))
                return (index, index); // if the text is fully contained within one run (unlikely)
            if (placeholder.Length < text.Length) continue; // realistically, this shouldn't happen

            foreach (var character in text)
            {
                if (character != placeholder[currentStop])
                {
                    currentStart = -1;
                    currentEnd = -1;
                    break;
                }

                if (currentStop == placeholder.Length - 1)
                {
                    currentEnd = index;
                    break;
                }

                currentStart = currentStart == -1 ? index : currentStart;
                currentStop++;
            }

            if (currentEnd != -1) break;
        }

        return (currentStart, currentEnd);
    }

    [GeneratedRegex(@"\{\{\s*([a-zA-Z0-9\-_]+)\s*(?::([^{}]*))?\s*\}\}")]
    private static partial Regex TextPlaceholderRegex();

    [GeneratedRegex(@"\[\[\s*([a-zA-Z0-9\-_]+)\s*(?::([^{}]*))?\s*\]\]")]
    private static partial Regex ListPlaceholderRegex();
}