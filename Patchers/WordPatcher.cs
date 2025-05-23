using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.JavaScript;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SimpleOfficePatchers.Extensions;
using SimpleOfficePatchers.Models.Word;

namespace SimpleOfficePatchers.Patchers;

public partial class WordPatcher
{
    [JSExport]
    public static byte[] PatchDocument(byte[] documentBytes, string serializedPatches)
    {
        var patches =
            JsonSerializer.Deserialize(serializedPatches, WordPatchContext.Default.DictionaryStringWordPatch);
        ArgumentNullException.ThrowIfNull(patches);
        var stream = new MemoryStream();
        stream.Write(documentBytes);

        var document = WordprocessingDocument.Open(stream, true);
        foreach (var (key, patch) in patches)
            ApplyPatch(document, key, patch);

        document.Dispose();
        return stream.ToArray();
    }

    [JSExport]
    public static string ExtractPlaceholders(byte[] documentBytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        var body = document.MainDocumentPart?.Document.Body;
        ArgumentNullException.ThrowIfNull(body);

        var result = new List<WordPlaceholder>();
        var matches = PlaceholderRegex().Matches(body.InnerText).Select(MatchToPlaceholder)
            .GroupBy(m => m.Name).ToDictionary(grouping => grouping.Key, grouping => grouping.ToList());

        foreach (var (_, placeholders) in matches)
        {
            var orderedPlaceholders = placeholders.OrderBy(pm => pm.Attribute).ThenBy(pm => pm.Description).ToList();
            foreach (var placeholderMatch in orderedPlaceholders) ProcessPlaceholder(placeholderMatch);
        }

        return JsonSerializer.Serialize(result,
            WordPlaceholderContext.Default.ListWordPlaceholder);

        void ProcessPlaceholder(PlaceholderMatch match)
        {
            if (result.Any(p => p.Name == match.Name)) return;
            switch (match.Attribute)
            {
                case "list":
                {
                    result.Add(new WordPlaceholder(match.Raw, match.Name, match.Description, WordPlaceholderType.List));
                    return;
                }
                case "group-title":
                {
                    if (matches.Any(kv => kv.Value.Any(pm => pm.Attribute == "group-items")))
                        result.Add(new WordPlaceholder($"{{{{{match.Name}}}}}", match.Name, match.Description,
                            WordPlaceholderType.GroupedList));
                    else
                        throw new Exception(
                            "found a placeholder with @group-title but not an accompanying @group-items");
                    return;
                }
                case "group-items":
                {
                    if (matches.Any(kv => kv.Value.Any(pm => pm.Attribute == "group-title")))
                        result.Add(new WordPlaceholder($"{{{{{match.Name}}}}}", match.Name, match.Description,
                            WordPlaceholderType.GroupedList));
                    else
                        throw new Exception(
                            "found a placeholder with @group-items but not an accompanying @group-title");
                    return;
                }
                case "":
                {
                    result.Add(new WordPlaceholder(match.Raw, match.Name, match.Description, WordPlaceholderType.Text));
                    return;
                }
                default:
                {
                    throw new Exception($"unknown attribute \"@{match.Attribute}\"");
                }
            }
        }
    }


    private static void ApplyPatch(WordprocessingDocument document, string key, WordPatch patch)
    {
        switch (patch.Type)
        {
            case WordPlaceholderType.Text:
            {
                var discriminatedPatch = (WordTextPatch)patch;
                ApplyTextPatch(document, FindParagraphsWithPlaceholder(document, key), discriminatedPatch);
                break;
            }
            case WordPlaceholderType.List:
            {
                var discriminatedPatch = (WordListPatch)patch;
                ApplyListPatch(document, FindParagraphsWithPlaceholder(document, key, "list"), discriminatedPatch);
                break;
            }
            case WordPlaceholderType.GroupedList:
            {
                var discriminatedPatch = (WordGroupListPatch)patch;
                ApplyGroupedListPatch(document, FindParagraphsWithPlaceholder(document, key, "group-title"),
                    FindParagraphsWithPlaceholder(document, key, "group-items"), discriminatedPatch);
                break;
            }
            default:
                throw new ArgumentOutOfRangeException();
        }
    }

    private static void ApplyGroupedListPatch(WordprocessingDocument document,
        List<PlaceholderInformation> titlePlaceholders, List<PlaceholderInformation> itemsPlaceholders,
        WordGroupListPatch patch)
    {
        if (patch.Groups.Count == 0) return;
        ArgumentNullException.ThrowIfNull(titlePlaceholders);
        ArgumentNullException.ThrowIfNull(itemsPlaceholders);
        if (titlePlaceholders.Count != itemsPlaceholders.Count)
            throw new Exception(
                $"the amount of @group-title and @group-items placeholders must match! (got {titlePlaceholders.Count} =/= ${itemsPlaceholders.Count})");

        var body = document.GetBodyOrThrow();
        for (var placeholdersIndex = 0; placeholdersIndex < titlePlaceholders.Count; placeholdersIndex++)
        {
            var titlePlaceholder = titlePlaceholders[placeholdersIndex];
            var itemsPlaceholder = itemsPlaceholders[placeholdersIndex];

            var parent = LaterParagraph(body, titlePlaceholder.SourceParagraph, itemsPlaceholder.SourceParagraph);
            for (var groupIndex = patch.Groups.Count - 1; groupIndex >= 0; groupIndex--)
            {
                var group = patch.Groups[groupIndex];
                var itemsIdentifier = TemporaryGroupedListPlaceholder("items", groupIndex);
                var titleIdentifier = TemporaryGroupedListPlaceholder("title", groupIndex);

                body.InsertAfter(ReplaceParagraphText(itemsPlaceholder, $"{{{{@list {itemsIdentifier}}}}}"), parent);
                body.InsertAfter(ReplaceParagraphText(titlePlaceholder, $"{{{{{titleIdentifier}}}}}"), parent);

                ApplyListPatch(document, FindParagraphsWithPlaceholder(document, itemsIdentifier, "list"),
                    new WordListPatch(null, group.Items));
                ApplyTextPatch(document, FindParagraphsWithPlaceholder(document, titleIdentifier), group.Title);
            }

            body.RemoveChild(titlePlaceholder.SourceParagraph);
            body.RemoveChild(itemsPlaceholder.SourceParagraph);
        }

        return;

        string TemporaryGroupedListPlaceholder(string purpose, int index)
        {
            return $"__grouped_list_${index}_{purpose}__";
        }
    }

    private static void ApplyListPatch(WordprocessingDocument document,
        List<PlaceholderInformation> placeholdersInformation,
        WordListPatch patch)
    {
        if (patch.Items.Count == 0) return;
        ArgumentNullException.ThrowIfNull(placeholdersInformation);

        var body = document.GetBodyOrThrow();
        foreach (var placeholderInformation in placeholdersInformation)
        {
            var parent = placeholderInformation.SourceParagraph;
            var numbering = CloneNumbering(document, parent);
            ArgumentNullException.ThrowIfNull(numbering); // TODO: provide more information

            foreach (var (index, item) in patch.Items.AsEnumerable().Reverse().Index())
                switch (item.Type)
                {
                    case WordPlaceholderType.Text:
                    {
                        var discriminatedItem = (WordTextPatch)item;

                        var newParagraph =
                            ReplaceParagraphText(placeholderInformation, discriminatedItem.Text, numbering);
                        body.InsertAfter(newParagraph, parent);
                        break;
                    }
                    case WordPlaceholderType.List:
                    {
                        var discriminatedItem = (WordListPatch)item;

                        var listIdentifier = TemporaryListPlaceholder(index);
                        body.InsertAfter(ReplaceParagraphText(placeholderInformation,
                            $"{{{{@list {listIdentifier}}}}}", numbering with { Level = numbering.Level + 1 }), parent);
                        ApplyPatch(document, listIdentifier, discriminatedItem);

                        if (discriminatedItem.SublistTitle != null)
                        {
                            var sublistTitleParagraph = ReplaceParagraphText(placeholderInformation,
                                discriminatedItem.SublistTitle.Text, numbering);
                            body.InsertAfter(sublistTitleParagraph, parent);
                        }

                        break;
                    }
                    case WordPlaceholderType.GroupedList:
                    {
                        var discriminatedItem = (WordGroupListPatch)item;

                        var listIdentifier = TemporaryListPlaceholder(index);
                        body.InsertAfter(
                            ReplaceParagraphText(placeholderInformation, $"{{{{@group-title {listIdentifier}}}}}",
                                numbering),
                            parent);
                        body.InsertAfter(ReplaceParagraphText(placeholderInformation,
                                $"{{{{@group-items {listIdentifier}}}}}",
                                numbering with { Level = numbering.Level + 1 }),
                            parent);
                        ApplyPatch(document, listIdentifier, discriminatedItem);
                        break;
                    }
                    default:
                        throw new ArgumentOutOfRangeException();
                }

            body.RemoveChild(parent);
        }

        return;

        string TemporaryListPlaceholder(int index)
        {
            return $"__list_${index}__";
        }
    }

    private static void ApplyTextPatch(WordprocessingDocument document,
        List<PlaceholderInformation> placeholdersInformation,
        WordTextPatch patch)
    {
        ArgumentNullException.ThrowIfNull(placeholdersInformation);

        foreach (var placeholderInformation in placeholdersInformation)
        {
            var newParagraph = ReplaceParagraphText(placeholderInformation, patch.Text);
            document.GetBodyOrThrow().ReplaceChild(newParagraph, placeholderInformation.SourceParagraph);
        }
    }

    private static List<PlaceholderInformation> FindParagraphsWithPlaceholder(WordprocessingDocument document,
        string name, string attribute = "")
    {
        var body = document.GetBodyOrThrow();
        var result = new List<PlaceholderInformation>();
        foreach (var child in body.ChildElements)
        {
            if (child is not Paragraph paragraph || !PlaceholderRegex().IsMatch(paragraph.InnerText)) continue;

            var exactMatch = PlaceholderRegex().Matches(paragraph.InnerText).Select(MatchToPlaceholder)
                .FirstOrDefault(m => m.Name == name && m.Attribute == attribute);
            if (exactMatch != null)
                result.Add(new PlaceholderInformation(FindRunIndices(paragraph, exactMatch.Raw), paragraph));
        }

        return result;
    }

    private static Paragraph ReplaceParagraphText(PlaceholderInformation placeholderInformation, string with,
        NumberingOptions numberingOptions = null)
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

        if (numberingOptions != null)
        {
            ArgumentNullException.ThrowIfNull(paragraph.ParagraphProperties);
            paragraph.ParagraphProperties.NumberingProperties = new NumberingProperties
            {
                NumberingId = new NumberingId { Val = numberingOptions.Id },
                NumberingLevelReference = new NumberingLevelReference { Val = numberingOptions.Level }
            };
        }

        return paragraph;
    }

    private static NumberingOptions CloneNumbering(WordprocessingDocument document, Paragraph source)
    {
        // TODO: perhaps don't throw on every failure 😭
        var id = source.ParagraphProperties?.NumberingProperties?.NumberingId?.Val;
        var level = source.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val ?? new Int32Value(0);
        if (id == null) return null;

        var numberingDefinitions = document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        ArgumentNullException.ThrowIfNull(numberingDefinitions);

        var numberingInstance = numberingDefinitions.Descendants<NumberingInstance>()
            .FirstOrDefault(ni => ni.NumberID?.Value == id);
        ArgumentNullException.ThrowIfNull(numberingInstance);

        var abstractNumberingInstance = numberingDefinitions.Descendants<AbstractNum>()
            .FirstOrDefault(ani => ani.AbstractNumberId == numberingInstance.AbstractNumId?.Val);
        ArgumentNullException.ThrowIfNull(abstractNumberingInstance);
        var startValue = abstractNumberingInstance.Descendants<Level>()
            .FirstOrDefault(lvl => lvl.LevelIndex == level)
            ?.StartNumberingValue?.Val ?? new Int32Value(1);

        var newInstance = (NumberingInstance)numberingInstance.Clone();
        newInstance.NumberID = new Int32Value(numberingDefinitions.Descendants<NumberingInstance>().Count() + 1);
        newInstance.AppendChild(new LevelOverride
        {
            StartOverrideNumberingValue = new StartOverrideNumberingValue { Val = startValue }
        });
        numberingDefinitions.AppendChild(newInstance);

        return new NumberingOptions(newInstance.NumberID.Value, level.Value);
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

    private static Paragraph LaterParagraph(Body body, Paragraph left, Paragraph right)
    {
        var paragraphs = body.Descendants<Paragraph>().ToList();
        return paragraphs.IndexOf(left) > paragraphs.IndexOf(right) ? left : right;
    }

    private static PlaceholderMatch MatchToPlaceholder(Match match)
    {
        if (!match.Success)
            throw new Exception(
                "MatchToPlaceholder must be called after verifying that the string matches the regex");
        return new PlaceholderMatch(match.Groups["attribute"].Value, match.Value, match.Groups["name"].Value,
            match.Groups["description"].Value);
    }

    [GeneratedRegex(@"\{\{\s*(?:@(?<attribute>[\w-]+)\s+)?(?<name>[\w_\-$@]+)\s*(?::(?<description>[^{}]*))?\s*\}\}")]
    private static partial Regex PlaceholderRegex();

    private record PlaceholderMatch(string Attribute, string Raw, string Name, string Description);

    private record PlaceholderInformation((int, int) RunIndices, Paragraph SourceParagraph);

    private record NumberingOptions(int Id, int Level);
}