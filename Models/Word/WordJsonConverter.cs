using System;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SimpleOfficePatchers.Models.Word;

public class WordPatchJsonConverter : JsonConverter<WordPatch>
{
    public override WordPatch Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType != JsonTokenType.StartObject) throw new JsonException();
        using var doc = JsonDocument.ParseValue(ref reader);
        if (!Enum.TryParse<WordPlaceholderType>(doc.RootElement.GetProperty("type").GetString(), out var type))
            throw new JsonException();

        switch (type)
        {
            case WordPlaceholderType.Text:
            {
                var text = doc.RootElement.GetProperty("text").GetString();
                if (text == null) throw new JsonException();
                return new WordTextPatch(text);
            }
            case WordPlaceholderType.List:
            {
                var items = doc.RootElement.GetProperty("items").EnumerateObject()
                    .ToDictionary(obj => obj.Name,
                        obj => obj.Value.EnumerateArray().Select(obj2 =>
                                JsonSerializer.Deserialize(obj2.GetRawText(), WordPatchContext.Default.WordPatch))
                            .ToList());
                return new WordListPatch(items);
            }
            case WordPlaceholderType.Group:
            {
                var groups = doc.RootElement.GetProperty("groups").EnumerateArray()
                    .Select(obj =>
                    {
                        var title = JsonSerializer.Deserialize(obj.GetProperty("title").GetRawText(),
                            WordPatchContext.Default.WordPatch);
                        var items = obj.GetProperty("items").EnumerateObject()
                            .ToDictionary(obj2 => obj2.Name,
                                obj2 => JsonSerializer.Deserialize(obj2.Value.GetRawText(),
                                    WordPatchContext.Default.WordPatch));
                        if (title.Type != WordPlaceholderType.Text) throw new JsonException();
                        return new WordGroupItem((WordTextPatch)title, items);
                    }).ToList();
                return new WordGroupPatch(groups);
            }
            default:
                throw new ArgumentOutOfRangeException();
        }
    }

    public override void Write(Utf8JsonWriter writer, WordPatch value, JsonSerializerOptions options)
    {
        writer.WriteStartObject();
        writer.WriteString("type", value.Type.ToString());
        switch (value.Type)
        {
            case WordPlaceholderType.Text:
            {
                var discriminatedPatch = (WordTextPatch)value;
                writer.WriteString("text", discriminatedPatch.Text);
                break;
            }
            case WordPlaceholderType.List:
            {
                var discriminatedPatch = (WordListPatch)value;

                writer.WriteStartArray("items");
                foreach (var item in discriminatedPatch.Items)
                    JsonSerializer.Serialize(writer, item, WordPatchContext.Default.WordPatch);
                writer.WriteEndArray();
                break;
            }
            case WordPlaceholderType.Group:
            {
                var discriminatedPatch = (WordGroupPatch)value;

                writer.WriteStartArray("groups");
                foreach (var group in discriminatedPatch.Groups)
                {
                    writer.WriteStartObject();
                    writer.WritePropertyName("title");
                    JsonSerializer.Serialize(writer, group.Title, WordPatchContext.Default.WordPatch);
                    writer.WriteStartArray("items");
                    foreach (var item in group.Items)
                        JsonSerializer.Serialize(writer, item, WordPatchContext.Default.WordPatch);
                    writer.WriteEndArray();
                    writer.WriteEndObject();
                }

                writer.WriteEndArray();
                break;
            }
            default:
                throw new ArgumentOutOfRangeException();
        }

        writer.WriteEndObject();
    }
}