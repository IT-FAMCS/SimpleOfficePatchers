using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace SimpleOfficePatchers.Models;

public record WordPlaceholder(string Raw, string Name, string Description, bool List);

[JsonSourceGenerationOptions(WriteIndented = true, PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(List<WordPlaceholder>))]
internal partial class WordPlaceholderContext : JsonSerializerContext;