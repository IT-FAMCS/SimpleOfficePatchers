using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SimpleOfficePatchers.Models.Word;

public enum WordPlaceholderType
{
    Text,
    List,
    Group
}

public record WordPlaceholder(
    string Raw,
    string Name,
    string Description,
    WordPlaceholderType Type,
    List<WordPlaceholder> Children);

[JsonSourceGenerationOptions(JsonSerializerDefaults.Web, UseStringEnumConverter = true)]
[JsonSerializable(typeof(List<WordPlaceholder>))]
internal partial class WordPlaceholderContext : JsonSerializerContext;