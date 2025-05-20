#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.JavaScript;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleOfficePatchers.Models;

namespace SimpleOfficePatchers.Patchers;

public partial class WordPatcher
{
    [JSExport]
    public static byte[] PatchDocument(byte[] bytes, string serializedPatches)
    {
        var patches =
            JsonSerializer.Deserialize(serializedPatches, WordPatchesContext.Default.WordPatches);
        ArgumentNullException.ThrowIfNull(patches);
        var stream = new MemoryStream();
        stream.Write(bytes);
        var document = WordprocessingDocument.Open(stream, true);

        foreach (var (placeholder, patch) in patches.Text)
            ApplyTextPatch(document,
                FindParagraphsWithPlaceholder(document, placeholder), patch);
        foreach (var (placeholder, patch) in patches.List)
            ApplyListPatch(document, FindParagraphsWithPlaceholder(document, placeholder), patch);

        document.Dispose();
        return stream.ToArray();
    }

    private record PlaceholderInformation((int, int) RunIndices, Paragraph SourceParagraph);

    private static void ApplyListPatch(WordprocessingDocument document,
        List<PlaceholderInformation>? placeholderInformations,
        WordListPatch patch)
    {
        if (patch.Patches.Count == 0) return;
        var body = document.MainDocumentPart?.Document.Body;
        ArgumentNullException.ThrowIfNull(placeholderInformations);
        ArgumentNullException.ThrowIfNull(body);

        foreach (var placeholderInformation in placeholderInformations)
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
        List<PlaceholderInformation>? placeholderInformations,
        WordTextPatch patch)
    {
        var body = document.MainDocumentPart?.Document.Body;
        ArgumentNullException.ThrowIfNull(placeholderInformations);
        ArgumentNullException.ThrowIfNull(body);

        foreach (var placeholderInformation in placeholderInformations)
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
            var completePlaceholder = $"{{{{{placeholder}}}}}";
            if (child is Paragraph paragraph && paragraph.InnerText.Contains(completePlaceholder))
                result.Add(new PlaceholderInformation(FindRunIndices(paragraph, completePlaceholder), paragraph));
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
}