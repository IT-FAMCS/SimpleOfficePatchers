using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.JavaScript;
using System.Text;
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
        foreach (var patch in patches)
            ApplyPatch(document, patch);

        document.Dispose();
        return stream.ToArray();
    }

    [JSExport]
    public static string ExtractPlaceholders(byte[] documentBytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(documentBytes), false);
        var body = document.MainDocumentPart?.Document.Body;
        ArgumentNullException.ThrowIfNull(body);

        var matches = PlaceholderRegex().Matches(body.InnerText).Select(MatchToPlaceholder).ToList();
        var result = matches.Select(ProcessPlaceholder).Where(pl =>
                matches.Where(m => m.Name == pl.Name).All(m => string.IsNullOrEmpty(m.Parent)))
            .ToList();

        return JsonSerializer.Serialize(result,
            WordPlaceholderContext.Default.ListWordPlaceholder);

        WordPlaceholder ProcessPlaceholder(PlaceholderMatch match)
        {
            var children = matches.Where(m => m.Parent == match.Name).Select(ProcessPlaceholder).ToList();
            switch (match.Attribute)
            {
                case "list":
                {
                    return new WordPlaceholder(match.ToString(), match.Name, match.Description,
                        WordPlaceholderType.List, children
                    );
                }
                case "group":
                {
                    return new WordPlaceholder(match.ToString(), match.Name, match.Description,
                        WordPlaceholderType.Group,
                        children);
                }
                case "":
                {
                    if (children.Count != 0) throw new Exception("a text placeholder cannot have children");
                    return new WordPlaceholder(match.ToString(), match.Name, match.Description,
                        WordPlaceholderType.Text, children
                    );
                }
                default:
                {
                    throw new Exception($"unknown attribute \"@{match.Attribute}\"");
                }
            }
        }
    }

    private static void ApplyPatch(WordprocessingDocument document, KeyValuePair<string, WordPatch> patch)
    {
        switch (patch.Value.Type)
        {
            case WordPlaceholderType.Text:
            {
                var paragraphs = FindParagraphsWithPlaceholder(document, patch.Key);
                foreach (var paragraph in paragraphs)
                    ApplyTextPatch(document, new PatchMetadata(patch.Key, patch.Value, paragraph));
                break;
            }
            case WordPlaceholderType.List:
            {
                var paragraphs = FindParagraphsWithPlaceholder(document, patch.Key, "list");
                foreach (var metadata in paragraphs.SelectMany(paragraph => CopyPatch(document,
                             new PatchMetadata(patch.Key, patch.Value, paragraph),
                             null,
                             paragraph.SourceParagraph, true)))
                    ApplyListPatch(document, metadata);

                break;
            }
            case WordPlaceholderType.Group:
            {
                var paragraphs = FindParagraphsWithPlaceholder(document, patch.Key, "group");
                foreach (var metadata in paragraphs.SelectMany(paragraph => CopyPatch(document,
                             new PatchMetadata(patch.Key, patch.Value, paragraph),
                             null,
                             paragraph.SourceParagraph, true)))
                    ApplyGroupPatch(document, metadata);

                break;
            }
            default:
                throw new ArgumentOutOfRangeException();
        }

        foreach (var paragraph in GetAllParagraphs(document, patch).Distinct())
            paragraph.Remove();
    }

    private static void ApplyTextPatch(WordprocessingDocument document, PatchMetadata metadata)
    {
        var text = ((WordTextPatch)metadata.Patch).Text;
        var body = document.GetBodyOrThrow();

        (metadata.Information.SourceParagraph.Parent ?? body).ReplaceChild(
            ReplaceParagraphText(metadata.Information, text),
            metadata.Information.SourceParagraph);
    }

    private static void ApplyListPatch(WordprocessingDocument document, PatchMetadata metadata)
    {
        var listPatch = (WordListPatch)metadata.Patch;
        var length = listPatch.Items.Values.First().Count;
        for (var index = 0; index < length; index++)
        {
            foreach (var key in listPatch.Items.Keys)
            {
                var item = listPatch.Items[key][index];
                var meta = (PatchMetadata)item.Meta;
                switch (item.Type)
                {
                    case WordPlaceholderType.Text: ApplyTextPatch(document, meta); break;
                    case WordPlaceholderType.List: ApplyListPatch(document, meta); break;
                    case WordPlaceholderType.Group: ApplyGroupPatch(document, meta); break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }

        metadata.Information.SourceParagraph.Remove();
    }

    private static void ApplyGroupPatch(WordprocessingDocument document, PatchMetadata metadata)
    {
        var groupPatch = (WordGroupPatch)metadata.Patch;
        var group = groupPatch.Groups.First();
        var body = document.GetBodyOrThrow();

        var title = ReplaceParagraphText(metadata.Information, group.Title.Text);
        (metadata.Information.SourceParagraph.Parent ?? body).ReplaceChild(title, metadata.Information.SourceParagraph);
        if (string.IsNullOrEmpty(group.Title.Text)) title.Remove();

        foreach (var (_, item) in group.Items)
        {
            var meta = (PatchMetadata)item.Meta;
            switch (item.Type)
            {
                case WordPlaceholderType.Text: ApplyTextPatch(document, meta); break;
                case WordPlaceholderType.List: ApplyListPatch(document, meta); break;
                case WordPlaceholderType.Group: ApplyGroupPatch(document, meta); break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
    }

    private record PatchMetadata(string Key, WordPatch Patch, PlaceholderInformation Information);

    private static List<Paragraph> GetAllParagraphs(WordprocessingDocument document,
        KeyValuePair<string, WordPatch> patch)
    {
        var result = new List<Paragraph>();
        result.AddRange(FindParagraphsWithPlaceholder(document, patch.Key).Select(pi => pi.SourceParagraph));
        switch (patch.Value.Type)
        {
            case WordPlaceholderType.Text:
                return result;
            case WordPlaceholderType.List:
            {
                result.AddRange(((WordListPatch)patch.Value).Items.SelectMany(p =>
                    p.Value.SelectMany(v =>
                        GetAllParagraphs(document, new KeyValuePair<string, WordPatch>(p.Key, v)))));
                return result;
            }
            case WordPlaceholderType.Group:
            {
                result.AddRange(((WordGroupPatch)patch.Value).Groups.SelectMany(g =>
                    g.Items.SelectMany(p =>
                        GetAllParagraphs(document, p))));
                return result;
            }
            default:
                throw new ArgumentOutOfRangeException();
        }
    }

    private static List<PatchMetadata> CopyPatch(WordprocessingDocument document,
        PatchMetadata patch,
        string parent,
        Paragraph where,
        bool before, string suffix = "_copy", NumberingOptions numberingOptions = null)
    {
        var body = document.GetBodyOrThrow();

        PatchMetadata GetPatchMetadata(string overrideSuffix = "")
        {
            var newKey = patch.Information.Match.Name +
                         (string.IsNullOrEmpty(overrideSuffix) ? suffix : overrideSuffix);
            var updatedMatch = patch.Information.Match with { Name = newKey, Parent = parent };
            var newParagraph = before
                ? body.InsertBefore(ReplaceParagraphText(patch.Information, updatedMatch.ToString(), numberingOptions),
                    where)
                : body.InsertAfter(ReplaceParagraphText(patch.Information, updatedMatch.ToString(), numberingOptions),
                    where);
            return patch with
            {
                Key = newKey,
                Information = new PlaceholderInformation(Match: updatedMatch,
                    RunIndices: FindRunIndices(newParagraph, updatedMatch.ToString()),
                    SourceParagraph: newParagraph),
            };
        }

        switch (patch.Patch.Type)
        {
            case WordPlaceholderType.Text: return [GetPatchMetadata()];
            case WordPlaceholderType.List:
            {
                var listPatch = (WordListPatch)patch.Patch;
                var patchMetadata = GetPatchMetadata();
                var listNumberingOptions = CloneNumbering(document, patchMetadata.Information.SourceParagraph);
                var copiedItems = new Dictionary<string, List<WordPatch>>();
                for (var index = listPatch.Items.Values.First().Count - 1; index >= 0; index--)
                {
                    var listSuffix = "_copy_list_" + index;
                    foreach (var key in listPatch.Items.Keys)
                    {
                        if (!copiedItems.ContainsKey(key)) copiedItems[key] = [];
                        var paragraphs = FindParagraphsWithPlaceholder(document, key);
                        foreach (var paragraph in paragraphs.AsEnumerable().Reverse())
                        {
                            var copies = CopyPatch(document,
                                new PatchMetadata(key, listPatch.Items[key][index], paragraph),
                                patchMetadata.Key,
                                patchMetadata.Information.SourceParagraph,
                                false, listSuffix, listNumberingOptions);
                            foreach (var copy in copies) copiedItems[key].Add(copy.Patch with { Meta = copy });
                        }
                    }
                }

                return [patchMetadata with { Patch = listPatch with { Items = copiedItems } }];
            }
            case WordPlaceholderType.Group:
            {
                var groupPatch = (WordGroupPatch)patch.Patch;
                var patchMetadatas = new List<PatchMetadata>();
                foreach (var (index, group) in groupPatch.Groups.AsEnumerable().Index())
                {
                    var groupSuffix = "_copy_group_" + index;
                    var patchMetadata = GetPatchMetadata(groupSuffix);
                    var copiedItems = new Dictionary<string, WordPatch>();
                    foreach (var (key, value) in group.Items)
                    {
                        var paragraphs = FindParagraphsWithPlaceholder(document, key);
                        foreach (var paragraph in paragraphs.AsEnumerable().Reverse())
                        {
                            var copies = CopyPatch(document, new PatchMetadata(key, value, paragraph),
                                patchMetadata.Key,
                                patchMetadata.Information.SourceParagraph,
                                false, groupSuffix);
                            foreach (var copy in copies) copiedItems[copy.Key] = copy.Patch with { Meta = copy };
                        }
                    }

                    patchMetadatas.Add(patchMetadata with
                    {
                        Patch = groupPatch with { Groups = [group with { Items = copiedItems }] }
                    });
                }

                return patchMetadatas;
            }
            default:
                throw new ArgumentOutOfRangeException();
        }
    }

    private static List<PlaceholderInformation> FindParagraphsWithPlaceholder(WordprocessingDocument document,
        string name, string attribute = null, string parent = null)
    {
        var body = document.GetBodyOrThrow();

        return (from paragraph in body.Descendants<Paragraph>()
                where PlaceholderRegex().IsMatch(paragraph.InnerText)
                let exactMatch = PlaceholderRegex()
                    .Matches(paragraph.InnerText)
                    .Select(MatchToPlaceholder)
                    .FirstOrDefault(m =>
                        m.Name == name && (parent == null || m.Parent == parent) &&
                        (attribute == null || m.Attribute == attribute))
                where exactMatch != null
                select new PlaceholderInformation(FindRunIndices(paragraph, exactMatch.ToString()), paragraph,
                    exactMatch))
            .ToList();
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
                paragraph.ChildElements[i]);
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


    private static PlaceholderMatch MatchToPlaceholder(Match match)
    {
        if (!match.Success)
            throw new Exception(
                "MatchToPlaceholder must be called after verifying that the string matches the regex");
        return new PlaceholderMatch(match.Groups["attribute"].Value, match.Groups["name"].Value,
            match.Groups["description"].Value, match.Groups["parent"].Value);
    }

    [GeneratedRegex(
        @"\{\{\s*(?:>(?<parent>[\w_\-$@]+)\s+)?\s*(?:@(?<attribute>[\w-]+)\s+)?(?<name>[\w_\-$@]+)\s*(?::(?<description>[^{}]*))?\s*\}\}")]
    private static partial Regex PlaceholderRegex();

    private record PlaceholderMatch(string Attribute, string Name, string Description, string Parent)
    {
        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("{{");
            if (!string.IsNullOrEmpty(Parent)) sb.Append($">{Parent} ");
            if (!string.IsNullOrEmpty(Attribute)) sb.Append($"@{Attribute} ");
            sb.Append(Name);
            if (!string.IsNullOrEmpty(Description)) sb.Append($": {Description}");
            sb.Append("}}");
            return sb.ToString();
        }
    }

    private record PlaceholderInformation((int, int) RunIndices, Paragraph SourceParagraph, PlaceholderMatch Match);

    private record NumberingOptions(int Id, int Level);
}