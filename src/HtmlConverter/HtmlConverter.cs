using System.Collections;
using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using HtmlAgilityPack;

namespace Antelcat.Wpf;

#if NET8_0_OR_GREATER
public static partial class HtmlConverter
#else
public static class HtmlConverter
#endif
{
    public static FlowDocument ConvertHtmlToFlowDocument(string html)
    {
        if (string.IsNullOrWhiteSpace(html)) return new FlowDocument();
        
        var htmlDocument = new HtmlDocument();
        htmlDocument.LoadHtml(html);

        return ConvertHtmlDocumentToFlowDocument(htmlDocument);
    }

    public static FlowDocument ConvertHtmlDocumentToFlowDocument(HtmlDocument htmlDocument)
    {
        var flowDocument = new FlowDocument();
        
        var bodyNode = htmlDocument.DocumentNode.SelectSingleNode("//body") ?? htmlDocument.DocumentNode;
        if (bodyNode == null) return flowDocument;

        foreach (var node in bodyNode.SelectNodes("//text()"))
        {
            node.InnerHtml = SpaceCharacterRegex().Replace(node.InnerHtml, " ");
        }

        var section = new Section { Name = "body" };
        flowDocument.Blocks.Add(section);
        foreach (var childNode in bodyNode.ChildNodes)
        {
            AddHtmlNodeToTextElement(childNode, section);
        }

        return flowDocument;
    }

    private static void AddHtmlNodeToTextElement(HtmlNode htmlNode, TextElement parent)
    {
        var textElement = CreateTextElement(htmlNode);
        if (textElement == null) return;
        AddTextElementToTextElement(textElement, parent);
        foreach (var child in htmlNode.ChildNodes) AddHtmlNodeToTextElement(child, textElement);
    }

    private static void AddTextElementToTextElement(TextElement child, TextElement parent)
    {
        ICollection? childCollection = parent switch
        {
            Section section => section.Blocks,
            Paragraph paragraph => paragraph.Inlines,
            AnchoredBlock anchoredBlock => anchoredBlock.Blocks,
            Span span => span.Inlines,
            ListItem listItem => listItem.Blocks,
            List list => list.ListItems,
            _ => null
        };
        switch (childCollection)
        {
            case BlockCollection blockCollection when child is Block block:
                blockCollection.Add(block);
                break;
            case BlockCollection blockCollection when child is Inline inline:
                blockCollection.Add(new Paragraph(inline));
                break;
            case InlineCollection inlineCollection when child is Inline inline:
                inlineCollection.Add(inline);
                break;
            case InlineCollection inlineCollection when child is Block block:
                inlineCollection.Add(new Figure(block));
                break;
            case ListItemCollection listItemCollection when child is ListItem listItem:
                listItemCollection.Add(listItem);
                break;
            case ListItemCollection listItemCollection when child is Block block:
                listItemCollection.Add(new ListItem { Blocks = { block } });
                break;
            case ListItemCollection listItemCollection when child is Inline inline:
                listItemCollection.Add(new ListItem { Blocks = { new Paragraph(inline) } });
                break;
            default:
                throw new InvalidOperationException();
        }
    }

    private static TextElement? CreateTextElement(HtmlNode htmlNode)
    {
        Span CreateSpan()
        {
            var span = new Span();
            if (TryConvertToColor(htmlNode.GetAttributeValue("color", null), out var color)) span.Foreground = new SolidColorBrush(color);
            if (double.TryParse(htmlNode.GetAttributeValue("size", null), out var size)) span.FontSize = size;
            if (htmlNode.GetAttributeValue("face", null) is { } face) span.FontFamily = new FontFamily(face);
            SetPropertyWhenNotNull(TextElement.FontWeightProperty,
                htmlNode.GetAttributeValue("weight", null) switch
                {
                    "bold" => FontWeights.Bold,
                    "normal" => FontWeights.Normal,
                    _ => null
                });
            SetPropertyWhenNotNull(TextElement.FontStyleProperty,
                htmlNode.GetAttributeValue("style", null) switch
                {
                    "italic" => FontStyles.Italic,
                    "normal" => FontStyles.Normal,
                    _ => null
                });
            SetPropertyWhenNotNull(Inline.TextDecorationsProperty,
                htmlNode.GetAttributeValue("decoration", null) switch
                {
                    "line-through" => TextDecorations.Strikethrough,
                    "overline" => TextDecorations.OverLine,
                    "baseline" => TextDecorations.Baseline,
                    "underline" => TextDecorations.Underline,
                    "none" => null,
                    _ => null
                });

            return span;

            bool TryConvertToColor(string value, out Color result)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    result = default;
                    return false;
                }

                try
                {
                    result = (Color)ColorConverter.ConvertFromString(value);
                    return true;
                }
                catch
                {
                    result = default;
                    return false;
                }
            }

            void SetPropertyWhenNotNull(DependencyProperty dp, object? value)
            {
                if (value != null) span.SetValue(dp, value);
            }
        }

        return htmlNode.NodeType switch
        {
            HtmlNodeType.Element => htmlNode.Name switch
            {
                "p" => new Paragraph(),
                "font" => CreateSpan(),
                "br" => new LineBreak(),
                "b" or "strong" => new Bold(),
                "i" or "em" => new Italic(),
                "ul" => new List { MarkerStyle = TextMarkerStyle.Disc },
                "ol" => new List { MarkerStyle = TextMarkerStyle.Decimal },
                "li" => new ListItem(),
                "a" => new Hyperlink
                {
                    NavigateUri = new Uri(htmlNode.GetAttributeValue("href", ""), UriKind.RelativeOrAbsolute)
                },
                "img" => new InlineUIContainer(new Image
                {
                    Source = new BitmapImage(new Uri(htmlNode.GetAttributeValue("src", null), UriKind.RelativeOrAbsolute)),
                    Width = double.TryParse(htmlNode.GetAttributeValue("width", null), out var value) ? value : double.NaN,
                    Height = double.TryParse(htmlNode.GetAttributeValue("height", null), out value) ? value : double.NaN,
                    Stretch = Stretch.Uniform
                }),
                "h1" or "h2" or "h3" or "h4" or "h5" or "h6" => new Paragraph
                {
                    FontSize = htmlNode.Name switch
                    {
                        "h1" => 32,
                        "h2" => 28,
                        "h3" => 24,
                        "h4" => 20,
                        "h5" => 16,
                        "h6" => 14,
                        _ => 12
                    },
                    FontWeight = FontWeights.Bold
                },
                _ => null
            },
            HtmlNodeType.Text when !string.IsNullOrWhiteSpace(htmlNode.InnerText) => new Run(WebUtility.HtmlDecode(htmlNode.InnerText)),
            _ => null
        };
    }

    public static string ConvertFlowDocumentToHtml(FlowDocument flowDocument)
    {
        var sb = new StringBuilder();
        sb.Append("<!doctype html><html><head><meta charset='UTF-8'></head><body>");
        foreach (var block in flowDocument.Blocks) ConvertBlockToHtml(block, sb);
        sb.Append("</body></html>");
        return sb.ToString();
    }

    private static void ConvertBlockToHtml(Block block, StringBuilder sb)
    {
        switch (block)
        {
            case Paragraph paragraph:
            {
                var tag = paragraph.ReadLocalValue(TextElement.FontSizeProperty) switch
                {
                    >= 32d => "h1",
                    >= 28d => "h2",
                    >= 24d => "h3",
                    >= 20d => "h4",
                    >= 16d => "h5",
                    >= 14d => "h6",
                    _ => "p"
                };
                sb.Append('<').Append(tag).Append('>');
                foreach (var inline in paragraph.Inlines)
                {
                    ConvertInlineToHtml(inline, sb);
                }
                sb.Append("</").Append(tag).Append('>');
                break;
            }
            case Section section:
            {
                foreach (var blockChild in section.Blocks)
                {
                    ConvertBlockToHtml(blockChild, sb);
                }
                break;
            }
            case List list:
            {
                sb.Append(list.MarkerStyle switch
                {
                    TextMarkerStyle.Disc => "<ul>",
                    TextMarkerStyle.Decimal => "<ol>",
                    _ => "<ul>"
                });
                foreach (var listItem in list.ListItems)
                {
                    sb.Append("<li>");
                    foreach (var inline in listItem.Blocks)
                    {
                        ConvertBlockToHtml(inline, sb);
                    }
                    sb.Append("</li>");
                }
                sb.Append(list.MarkerStyle switch
                {
                    TextMarkerStyle.Disc => "</ul>",
                    TextMarkerStyle.Decimal => "</ol>",
                    _ => "</ul>"
                });
                break;
            }
            case BlockUIContainer { Child: Image image }:
            {
                ConvertImageToHtml(image, sb);
                break;
            }
        }
    }

    private static void ConvertInlineToHtml(Inline inline, StringBuilder sb)
    {
        bool TryReadLocalValue(DependencyProperty dp, out object? value)
        {
            value = inline.ReadLocalValue(dp);
            return value != DependencyProperty.UnsetValue;
        }

        var isInsideFont = false;

        void WriteFontProperty(string property)
        {
            if (!isInsideFont)
            {
                sb.Append("<font");
                isInsideFont = true;
            }
            sb.Append(' ').Append(property);
        }

        if (TryReadLocalValue(TextElement.ForegroundProperty, out var foreground) && foreground is SolidColorBrush foregroundBrush)
        {
            WriteFontProperty($"color=\"{foregroundBrush.Color.ToString(CultureInfo.InvariantCulture)}\"");
        }
        if (TryReadLocalValue(TextElement.FontFamilyProperty, out var fontFamily) && fontFamily is FontFamily fontFamilyValue)
        {
            WriteFontProperty($"face=\"{fontFamilyValue.Source}\"");
        }
        if (TryReadLocalValue(TextElement.FontSizeProperty, out var fontSize) && fontSize is double fontSizeValue)
        {
            WriteFontProperty($"size=\"{fontSizeValue}\"");
        }
        if (TryReadLocalValue(TextElement.FontWeightProperty, out var fontWeight) && fontWeight is FontWeight fontWeightValue)
        {
            WriteFontProperty($"weight=\"{(fontWeightValue == FontWeights.Bold ? "bold" : "normal")}\"");
        }
        if (TryReadLocalValue(TextElement.FontStyleProperty, out var fontStyle) && fontStyle is FontStyle fontStyleValue)
        {
            WriteFontProperty($"style=\"{(fontStyleValue == FontStyles.Italic ? "italic" : "normal")}\"");
        }
        if (TryReadLocalValue(Inline.TextDecorationsProperty, out var textDecorations) &&
            textDecorations is TextDecorationCollection textDecorationsValue)
        {
            var decoration = textDecorationsValue.Count switch
            {
                0 => "none",
                1 => textDecorationsValue[0].Location switch
                {
                    TextDecorationLocation.Strikethrough => "line-through",
                    TextDecorationLocation.OverLine => "overline",
                    TextDecorationLocation.Baseline => "baseline",
                    TextDecorationLocation.Underline => "underline",
                    _ => "none"
                },
                _ => "none"
            };
            WriteFontProperty($"decoration=\"{decoration}\"");
        }

        if (isInsideFont) sb.Append('>');

        switch (inline)
        {
            case Run run:
            {
                sb.Append(WebUtility.HtmlEncode(run.Text));
                break;
            }
            case LineBreak:
            {
                sb.Append("<br>");
                break;
            }
            case Bold bold:
            {
                sb.Append("<b>");
                foreach (var inlineChild in bold.Inlines)
                {
                    ConvertInlineToHtml(inlineChild, sb);
                }
                sb.Append("</b>");
                break;
            }
            case Italic italic:
            {
                sb.Append("<i>");
                foreach (var inlineChild in italic.Inlines)
                {
                    ConvertInlineToHtml(inlineChild, sb);
                }
                sb.Append("</i>");
                break;
            }
            case Hyperlink hyperlink:
            {
                sb.Append($"<a href=\"{hyperlink.NavigateUri}\">");
                foreach (var inlineChild in hyperlink.Inlines)
                {
                    ConvertInlineToHtml(inlineChild, sb);
                }
                sb.Append("</a>");
                break;
            }
            case InlineUIContainer { Child: Image image }:
            {
                ConvertImageToHtml(image, sb);
                break;
            }
        }

        if (isInsideFont) sb.Append("</font>");
    }

    private static void ConvertImageToHtml(Image image, StringBuilder sb)
    {
        if (image.Source is not BitmapImage source) return;
        sb.Append($"<img src=\"{source.UriSource}\"");
        if (!double.IsNaN(image.Width)) sb.Append($" width=\"{image.Width}\"");
        if (!double.IsNaN(image.Height)) sb.Append($" height=\"{image.Height}\"");
        sb.Append('>');
    }

#if NET8_0_OR_GREATER
    [GeneratedRegex(@"\s+")]
    private static partial Regex SpaceCharacterRegex();
#else
    private static Regex SpaceCharacterRegex() => new(@"\s+");
#endif
}