using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SimpleOfficePatchers.Models.Word;

public abstract record WordPatch(WordPlaceholderType Type, [property: JsonIgnore] object Meta = null);

public record WordTextPatch(string Text) : WordPatch(WordPlaceholderType.Text);

public record WordListPatch(Dictionary<string, List<WordPatch>> Items)
    : WordPatch(WordPlaceholderType.List);

public record WordGroupItem(WordTextPatch Title, Dictionary<string, WordPatch> Items);

public record WordGroupPatch(List<WordGroupItem> Groups)
    : WordPatch(WordPlaceholderType.Group);

[JsonSourceGenerationOptions(JsonSerializerDefaults.Web, UseStringEnumConverter = true,
    Converters = [typeof(WordPatchJsonConverter)])]
[JsonSerializable(typeof(WordPatch))]
[JsonSerializable(typeof(Dictionary<string, WordPatch>))]
internal partial class WordPatchContext : JsonSerializerContext;