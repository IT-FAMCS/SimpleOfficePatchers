using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SimpleOfficePatchers.Extensions;

public static class WordprocessingDocumentExtensions
{
    public static Body GetBodyOrThrow(this WordprocessingDocument document)
    {
        var body = document.MainDocumentPart?.Document.Body;
        ArgumentNullException.ThrowIfNull(body);
        return body;
    }
}