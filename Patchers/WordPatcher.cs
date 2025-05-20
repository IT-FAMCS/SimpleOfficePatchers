#nullable enable
using System;
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
            ApplyTextPatch(ref document,
                FindParagraphWithPlaceholder(document, placeholder), patch);

        document.Dispose();
        return stream.ToArray();
    }

    private record PlaceholderInformation(int Location, (int, int) RunIndices, Paragraph SourceParagraph);

    private static void ApplyTextPatch(ref WordprocessingDocument document,
        PlaceholderInformation? placeholderInformation,
        WordTextPatch patch)
    {
        var body = document.MainDocumentPart?.Document.Body;
        ArgumentNullException.ThrowIfNull(placeholderInformation);
        ArgumentNullException.ThrowIfNull(body);

        var newParagraph = ReplaceParagraphText(placeholderInformation, patch.Text);
        body.ReplaceChild(newParagraph, placeholderInformation.SourceParagraph);
    }

    private static PlaceholderInformation? FindParagraphWithPlaceholder(WordprocessingDocument document,
        string placeholder)
    {
        var body = document.MainDocumentPart?.Document.Body;
        if (body is null) return null;

        foreach (var (index, child) in body.ChildElements.Index())
        {
            var completePlaceholder = $"{{{{{placeholder}}}}}";
            if (child is Paragraph paragraph && paragraph.InnerText.Contains(completePlaceholder))
                return new PlaceholderInformation(index, FindRunIndices(paragraph, completePlaceholder), paragraph);
        }

        return null;
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