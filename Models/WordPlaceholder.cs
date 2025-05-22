using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SimpleOfficePatchers.Models;

public enum WordPlaceholderType
{
    Text,
    List,
    GroupedList
}

public record WordPlaceholder(string Raw, string Name, string Description, WordPlaceholderType Type);

[JsonSourceGenerationOptions(JsonSerializerDefaults.Web)]
[JsonSerializable(typeof(List<WordPlaceholder>))]
internal partial class WordPlaceholderContext : JsonSerializerContext;