#nullable enable
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SimpleOfficePatchers.Models;

public record WordTextPatch(string Text);

public record WordListPatch(List<WordTextPatch> Items);

public record WordGroupListItem(WordTextPatch Title, WordListPatch List);

public record WordGroupListPatch(List<WordGroupListItem> Groups);

public class WordPatches
{
    public Dictionary<string, WordTextPatch>? Text { get; set; }
    public Dictionary<string, WordListPatch>? List { get; set; }
    public Dictionary<string, WordGroupListPatch>? GroupedList { get; set; }
}

[JsonSourceGenerationOptions(JsonSerializerDefaults.Web)]
[JsonSerializable(typeof(WordPatches))]
internal partial class WordPatchesContext : JsonSerializerContext;