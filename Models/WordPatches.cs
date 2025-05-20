using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace SimpleOfficePatchers.Models;

public record WordTextPatch(string Text);

public record WordListPatch(List<WordTextPatch> Patches);

public class WordPatches
{
    public Dictionary<string, WordTextPatch> Text { get; set; }
    public Dictionary<string, WordListPatch> List { get; set; }
}

[JsonSourceGenerationOptions(WriteIndented = true, PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(WordPatches))]
internal partial class WordPatchesContext : JsonSerializerContext;