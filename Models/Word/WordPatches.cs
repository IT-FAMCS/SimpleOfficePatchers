using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SimpleOfficePatchers.Models.Word;

public abstract record WordPatch(WordPlaceholderType Type);

public record WordTextPatch(string Text) : WordPatch(WordPlaceholderType.Text);

public record WordListPatch(WordTextPatch SublistTitle, List<WordPatch> Items)
    : WordPatch(WordPlaceholderType.List);

public record WordGroupListItem(WordTextPatch Title, List<WordPatch> Items);

public record WordGroupListPatch(List<WordGroupListItem> Groups)
    : WordPatch(WordPlaceholderType.GroupedList);

[JsonSourceGenerationOptions(JsonSerializerDefaults.Web, UseStringEnumConverter = true, Converters = [typeof(WordPatchJsonConverter)])]
[JsonSerializable(typeof(WordPatch))]
[JsonSerializable(typeof(Dictionary<string, WordPatch>))]
internal partial class WordPatchContext : JsonSerializerContext;