using System;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace PowerPointGenerator
{
    public class Template
    {
        // Creates a PresentationDocument.
        public static void CreatePackage(string filePath)
        {
            using (var package = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation)) {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private static void CreateParts(PresentationDocument document)
        {
            var thumbnailPart1 = document.AddNewPart<ThumbnailPart>("image/jpeg", "rId2");
            GenerateThumbnailPart1Content(thumbnailPart1);

            var presentationPart1 = document.AddPresentationPart();
            GeneratePresentationPart1Content(presentationPart1);

            var presentationPropertiesPart1 = presentationPart1.AddNewPart<PresentationPropertiesPart>("rId3");
            GeneratePresentationPropertiesPart1Content(presentationPropertiesPart1);

            var slidePart1 = presentationPart1.AddNewPart<SlidePart>("rId2");
            GenerateSlidePart1Content(slidePart1);

            var slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            GenerateSlideLayoutPart1Content(slideLayoutPart1);

            var slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
            GenerateSlideMasterPart1Content(slideMasterPart1);

            var slideLayoutPart2 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId8");
            GenerateSlideLayoutPart2Content(slideLayoutPart2);

            slideLayoutPart2.AddPart(slideMasterPart1, "rId1");

            var slideLayoutPart3 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId3");
            GenerateSlideLayoutPart3Content(slideLayoutPart3);

            slideLayoutPart3.AddPart(slideMasterPart1, "rId1");

            var slideLayoutPart4 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId7");
            GenerateSlideLayoutPart4Content(slideLayoutPart4);

            slideLayoutPart4.AddPart(slideMasterPart1, "rId1");

            var themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId12");
            GenerateThemePart1Content(themePart1);

            var slideLayoutPart5 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId2");
            GenerateSlideLayoutPart5Content(slideLayoutPart5);

            slideLayoutPart5.AddPart(slideMasterPart1, "rId1");

            var slideLayoutPart6 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId1");
            GenerateSlideLayoutPart6Content(slideLayoutPart6);

            slideLayoutPart6.AddPart(slideMasterPart1, "rId1");

            slideMasterPart1.AddPart(slideLayoutPart1, "rId6");

            var slideLayoutPart7 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId11");
            GenerateSlideLayoutPart7Content(slideLayoutPart7);

            slideLayoutPart7.AddPart(slideMasterPart1, "rId1");

            var slideLayoutPart8 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId5");
            GenerateSlideLayoutPart8Content(slideLayoutPart8);

            slideLayoutPart8.AddPart(slideMasterPart1, "rId1");

            var slideLayoutPart9 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId10");
            GenerateSlideLayoutPart9Content(slideLayoutPart9);

            slideLayoutPart9.AddPart(slideMasterPart1, "rId1");

            var slideLayoutPart10 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId4");
            GenerateSlideLayoutPart10Content(slideLayoutPart10);

            slideLayoutPart10.AddPart(slideMasterPart1, "rId1");

            var slideLayoutPart11 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId9");
            GenerateSlideLayoutPart11Content(slideLayoutPart11);

            slideLayoutPart11.AddPart(slideMasterPart1, "rId1");

            presentationPart1.AddPart(slideMasterPart1, "rId1");

            var tableStylesPart1 = presentationPart1.AddNewPart<TableStylesPart>("rId6");
            GenerateTableStylesPart1Content(tableStylesPart1);

            presentationPart1.AddPart(themePart1, "rId5");

            var viewPropertiesPart1 = presentationPart1.AddNewPart<ViewPropertiesPart>("rId4");
            GenerateViewPropertiesPart1Content(viewPropertiesPart1);

            var extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId4");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            SetPackageProperties(document);
        }

        // Generates content of thumbnailPart1.
        private static void GenerateThumbnailPart1Content(ThumbnailPart thumbnailPart1)
        {
            var data = GetBinaryDataStream(ThumbnailPart1Data);
            thumbnailPart1.FeedData(data);
            data.Close();
        }

        // Generates content of presentationPart1.
        private static void GeneratePresentationPart1Content(PresentationPart presentationPart1)
        {
            var presentation1 = new Presentation() {SaveSubsetFonts = true};
            presentation1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentation1.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentation1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var slideMasterIdList1 = new SlideMasterIdList();
            var slideMasterId1 = new SlideMasterId() {Id = (UInt32Value) 2147483648U, RelationshipId = "rId1"};

            slideMasterIdList1.Append(slideMasterId1);

            var slideIdList1 = new SlideIdList();
            var slideId1 = new SlideId() {Id = (UInt32Value) 256U, RelationshipId = "rId2"};

            slideIdList1.Append(slideId1);
            var slideSize1 = new SlideSize() {Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3};
            var notesSize1 = new NotesSize() {Cx = 6858000L, Cy = 9144000L};

            var defaultTextStyle1 = new DefaultTextStyle();

            var defaultParagraphProperties1 = new A.DefaultParagraphProperties();
            var defaultRunProperties1 = new A.DefaultRunProperties() {Language = "nl-NL"};

            defaultParagraphProperties1.Append(defaultRunProperties1);

            var level1ParagraphProperties1 = new A.Level1ParagraphProperties() {
                LeftMargin = 0,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties2 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill1 = new A.SolidFill();
            var schemeColor1 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill1.Append(schemeColor1);
            var latinFont1 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont1 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont1 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties2.Append(solidFill1);
            defaultRunProperties2.Append(latinFont1);
            defaultRunProperties2.Append(eastAsianFont1);
            defaultRunProperties2.Append(complexScriptFont1);

            level1ParagraphProperties1.Append(defaultRunProperties2);

            var level2ParagraphProperties1 = new A.Level2ParagraphProperties() {
                LeftMargin = 457200,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties3 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill2 = new A.SolidFill();
            var schemeColor2 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill2.Append(schemeColor2);
            var latinFont2 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont2 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont2 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties3.Append(solidFill2);
            defaultRunProperties3.Append(latinFont2);
            defaultRunProperties3.Append(eastAsianFont2);
            defaultRunProperties3.Append(complexScriptFont2);

            level2ParagraphProperties1.Append(defaultRunProperties3);

            var level3ParagraphProperties1 = new A.Level3ParagraphProperties() {
                LeftMargin = 914400,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties4 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill3 = new A.SolidFill();
            var schemeColor3 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill3.Append(schemeColor3);
            var latinFont3 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont3 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont3 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties4.Append(solidFill3);
            defaultRunProperties4.Append(latinFont3);
            defaultRunProperties4.Append(eastAsianFont3);
            defaultRunProperties4.Append(complexScriptFont3);

            level3ParagraphProperties1.Append(defaultRunProperties4);

            var level4ParagraphProperties1 = new A.Level4ParagraphProperties() {
                LeftMargin = 1371600,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties5 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill4 = new A.SolidFill();
            var schemeColor4 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill4.Append(schemeColor4);
            var latinFont4 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont4 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont4 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties5.Append(solidFill4);
            defaultRunProperties5.Append(latinFont4);
            defaultRunProperties5.Append(eastAsianFont4);
            defaultRunProperties5.Append(complexScriptFont4);

            level4ParagraphProperties1.Append(defaultRunProperties5);

            var level5ParagraphProperties1 = new A.Level5ParagraphProperties() {
                LeftMargin = 1828800,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties6 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill5 = new A.SolidFill();
            var schemeColor5 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill5.Append(schemeColor5);
            var latinFont5 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont5 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont5 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties6.Append(solidFill5);
            defaultRunProperties6.Append(latinFont5);
            defaultRunProperties6.Append(eastAsianFont5);
            defaultRunProperties6.Append(complexScriptFont5);

            level5ParagraphProperties1.Append(defaultRunProperties6);

            var level6ParagraphProperties1 = new A.Level6ParagraphProperties() {
                LeftMargin = 2286000,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties7 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill6 = new A.SolidFill();
            var schemeColor6 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill6.Append(schemeColor6);
            var latinFont6 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont6 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont6 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties7.Append(solidFill6);
            defaultRunProperties7.Append(latinFont6);
            defaultRunProperties7.Append(eastAsianFont6);
            defaultRunProperties7.Append(complexScriptFont6);

            level6ParagraphProperties1.Append(defaultRunProperties7);

            var level7ParagraphProperties1 = new A.Level7ParagraphProperties() {
                LeftMargin = 2743200,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties8 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill7 = new A.SolidFill();
            var schemeColor7 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill7.Append(schemeColor7);
            var latinFont7 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont7 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont7 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties8.Append(solidFill7);
            defaultRunProperties8.Append(latinFont7);
            defaultRunProperties8.Append(eastAsianFont7);
            defaultRunProperties8.Append(complexScriptFont7);

            level7ParagraphProperties1.Append(defaultRunProperties8);

            var level8ParagraphProperties1 = new A.Level8ParagraphProperties() {
                LeftMargin = 3200400,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties9 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill8 = new A.SolidFill();
            var schemeColor8 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill8.Append(schemeColor8);
            var latinFont8 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont8 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont8 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties9.Append(solidFill8);
            defaultRunProperties9.Append(latinFont8);
            defaultRunProperties9.Append(eastAsianFont8);
            defaultRunProperties9.Append(complexScriptFont8);

            level8ParagraphProperties1.Append(defaultRunProperties9);

            var level9ParagraphProperties1 = new A.Level9ParagraphProperties() {
                LeftMargin = 3657600,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties10 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill9 = new A.SolidFill();
            var schemeColor9 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill9.Append(schemeColor9);
            var latinFont9 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont9 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont9 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties10.Append(solidFill9);
            defaultRunProperties10.Append(latinFont9);
            defaultRunProperties10.Append(eastAsianFont9);
            defaultRunProperties10.Append(complexScriptFont9);

            level9ParagraphProperties1.Append(defaultRunProperties10);

            defaultTextStyle1.Append(defaultParagraphProperties1);
            defaultTextStyle1.Append(level1ParagraphProperties1);
            defaultTextStyle1.Append(level2ParagraphProperties1);
            defaultTextStyle1.Append(level3ParagraphProperties1);
            defaultTextStyle1.Append(level4ParagraphProperties1);
            defaultTextStyle1.Append(level5ParagraphProperties1);
            defaultTextStyle1.Append(level6ParagraphProperties1);
            defaultTextStyle1.Append(level7ParagraphProperties1);
            defaultTextStyle1.Append(level8ParagraphProperties1);
            defaultTextStyle1.Append(level9ParagraphProperties1);

            presentation1.Append(slideMasterIdList1);
            presentation1.Append(slideIdList1);
            presentation1.Append(slideSize1);
            presentation1.Append(notesSize1);
            presentation1.Append(defaultTextStyle1);

            presentationPart1.Presentation = presentation1;
        }

        // Generates content of presentationPropertiesPart1.
        private static void GeneratePresentationPropertiesPart1Content(PresentationPropertiesPart presentationPropertiesPart1)
        {
            var presentationProperties1 = new PresentationProperties();
            presentationProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentationProperties1.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentationProperties1.AddNamespaceDeclaration("p",
                "http://schemas.openxmlformats.org/presentationml/2006/main");

            var presentationPropertiesExtensionList1 = new PresentationPropertiesExtensionList();

            var presentationPropertiesExtension1 = new PresentationPropertiesExtension() {
                Uri = "{E76CE94A-603C-4142-B9EB-6D1370010A27}"
            };

            var discardImageEditData1 = new P14.DiscardImageEditData() {Val = false};
            discardImageEditData1.AddNamespaceDeclaration("p14",
                "http://schemas.microsoft.com/office/powerpoint/2010/main");

            presentationPropertiesExtension1.Append(discardImageEditData1);

            var presentationPropertiesExtension2 = new PresentationPropertiesExtension() {
                Uri = "{D31A062A-798A-4329-ABDD-BBA856620510}"
            };

            var defaultImageDpi1 = new P14.DefaultImageDpi() {Val = (UInt32Value) 220U};
            defaultImageDpi1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            presentationPropertiesExtension2.Append(defaultImageDpi1);

            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension1);
            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension2);

            presentationProperties1.Append(presentationPropertiesExtensionList1);

            presentationPropertiesPart1.PresentationProperties = presentationProperties1;
        }

        // Generates content of slidePart1.
        private static void GenerateSlidePart1Content(SlidePart slidePart1)
        {
            var slide1 = new Slide();
            slide1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData1 = new CommonSlideData();

            var shapeTree1 = new ShapeTree();

            var nonVisualGroupShapeProperties1 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties1 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties1 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(nonVisualGroupShapeDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(applicationNonVisualDrawingProperties1);

            var groupShapeProperties1 = new GroupShapeProperties();

            var transformGroup1 = new A.TransformGroup();
            var offset1 = new A.Offset() {X = 0L, Y = 0L};
            var extents1 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset1 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents1 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup1.Append(offset1);
            transformGroup1.Append(extents1);
            transformGroup1.Append(childOffset1);
            transformGroup1.Append(childExtents1);

            groupShapeProperties1.Append(transformGroup1);

            var shape1 = new Shape();

            var nonVisualShapeProperties1 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties2 = new NonVisualDrawingProperties() {Id = (UInt32Value) 4U, Name = "Title 3"};

            var nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties();
            var shapeLocks1 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties1.Append(shapeLocks1);

            var applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape1 = new PlaceholderShape() {Type = PlaceholderValues.Title};

            applicationNonVisualDrawingProperties2.Append(placeholderShape1);

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties2);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties2);
            var shapeProperties1 = new ShapeProperties();

            var textBody1 = new TextBody();
            var bodyProperties1 = new A.BodyProperties();
            var listStyle1 = new A.ListStyle();

            var paragraph1 = new A.Paragraph();
            var endParagraphRunProperties1 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(textBody1);

            shapeTree1.Append(nonVisualGroupShapeProperties1);
            shapeTree1.Append(groupShapeProperties1);
            shapeTree1.Append(shape1);

            var commonSlideDataExtensionList1 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension1 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId1 = new P14.CreationId() {Val = (UInt32Value) 20133129U};
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension1.Append(creationId1);

            commonSlideDataExtensionList1.Append(commonSlideDataExtension1);

            commonSlideData1.Append(shapeTree1);
            commonSlideData1.Append(commonSlideDataExtensionList1);

            var colorMapOverride1 = new ColorMapOverride();
            var masterColorMapping1 = new A.MasterColorMapping();

            colorMapOverride1.Append(masterColorMapping1);

            slide1.Append(commonSlideData1);
            slide1.Append(colorMapOverride1);

            slidePart1.Slide = slide1;
        }

        // Generates content of slideLayoutPart1.
        private static void GenerateSlideLayoutPart1Content(SlideLayoutPart slideLayoutPart1)
        {
            var slideLayout1 = new SlideLayout() {Type = SlideLayoutValues.TitleOnly, Preserve = true};
            slideLayout1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout1.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData2 = new CommonSlideData() {Name = "Title Only"};

            var shapeTree2 = new ShapeTree();

            var nonVisualGroupShapeProperties2 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties3 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties2 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties3 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties2.Append(nonVisualDrawingProperties3);
            nonVisualGroupShapeProperties2.Append(nonVisualGroupShapeDrawingProperties2);
            nonVisualGroupShapeProperties2.Append(applicationNonVisualDrawingProperties3);

            var groupShapeProperties2 = new GroupShapeProperties();

            var transformGroup2 = new A.TransformGroup();
            var offset2 = new A.Offset() {X = 0L, Y = 0L};
            var extents2 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset2 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents2 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup2.Append(offset2);
            transformGroup2.Append(extents2);
            transformGroup2.Append(childOffset2);
            transformGroup2.Append(childExtents2);

            groupShapeProperties2.Append(transformGroup2);

            var shape2 = new Shape();

            var nonVisualShapeProperties2 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties4 = new NonVisualDrawingProperties() {Id = (UInt32Value) 2U, Name = "Title 1"};

            var nonVisualShapeDrawingProperties2 = new NonVisualShapeDrawingProperties();
            var shapeLocks2 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties2.Append(shapeLocks2);

            var applicationNonVisualDrawingProperties4 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape2 = new PlaceholderShape() {Type = PlaceholderValues.Title};

            applicationNonVisualDrawingProperties4.Append(placeholderShape2);

            nonVisualShapeProperties2.Append(nonVisualDrawingProperties4);
            nonVisualShapeProperties2.Append(nonVisualShapeDrawingProperties2);
            nonVisualShapeProperties2.Append(applicationNonVisualDrawingProperties4);
            var shapeProperties2 = new ShapeProperties();

            var textBody2 = new TextBody();
            var bodyProperties2 = new A.BodyProperties();
            var listStyle2 = new A.ListStyle();

            var paragraph2 = new A.Paragraph();

            var run1 = new A.Run();

            var runProperties1 = new A.RunProperties() {Language = "en-US"};
            runProperties1.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text1 = new A.Text();
            text1.Text = "Click to edit Master title style";

            run1.Append(runProperties1);
            run1.Append(text1);
            var endParagraphRunProperties2 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph2.Append(run1);
            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);

            shape2.Append(nonVisualShapeProperties2);
            shape2.Append(shapeProperties2);
            shape2.Append(textBody2);

            var shape3 = new Shape();

            var nonVisualShapeProperties3 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties5 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Date Placeholder 2"
            };

            var nonVisualShapeDrawingProperties3 = new NonVisualShapeDrawingProperties();
            var shapeLocks3 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties3.Append(shapeLocks3);

            var applicationNonVisualDrawingProperties5 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape3 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 10U
            };

            applicationNonVisualDrawingProperties5.Append(placeholderShape3);

            nonVisualShapeProperties3.Append(nonVisualDrawingProperties5);
            nonVisualShapeProperties3.Append(nonVisualShapeDrawingProperties3);
            nonVisualShapeProperties3.Append(applicationNonVisualDrawingProperties5);
            var shapeProperties3 = new ShapeProperties();

            var textBody3 = new TextBody();
            var bodyProperties3 = new A.BodyProperties();
            var listStyle3 = new A.ListStyle();

            var paragraph3 = new A.Paragraph();

            var field1 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties2 = new A.RunProperties() {Language = "en-GB"};
            runProperties2.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text2 = new A.Text();
            text2.Text = "04/06/2014";

            field1.Append(runProperties2);
            field1.Append(text2);
            var endParagraphRunProperties3 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph3.Append(field1);
            paragraph3.Append(endParagraphRunProperties3);

            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph3);

            shape3.Append(nonVisualShapeProperties3);
            shape3.Append(shapeProperties3);
            shape3.Append(textBody3);

            var shape4 = new Shape();

            var nonVisualShapeProperties4 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties6 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Footer Placeholder 3"
            };

            var nonVisualShapeDrawingProperties4 = new NonVisualShapeDrawingProperties();
            var shapeLocks4 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties4.Append(shapeLocks4);

            var applicationNonVisualDrawingProperties6 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape4 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 11U
            };

            applicationNonVisualDrawingProperties6.Append(placeholderShape4);

            nonVisualShapeProperties4.Append(nonVisualDrawingProperties6);
            nonVisualShapeProperties4.Append(nonVisualShapeDrawingProperties4);
            nonVisualShapeProperties4.Append(applicationNonVisualDrawingProperties6);
            var shapeProperties4 = new ShapeProperties();

            var textBody4 = new TextBody();
            var bodyProperties4 = new A.BodyProperties();
            var listStyle4 = new A.ListStyle();

            var paragraph4 = new A.Paragraph();
            var endParagraphRunProperties4 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph4.Append(endParagraphRunProperties4);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph4);

            shape4.Append(nonVisualShapeProperties4);
            shape4.Append(shapeProperties4);
            shape4.Append(textBody4);

            var shape5 = new Shape();

            var nonVisualShapeProperties5 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties7 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 5U,
                Name = "Slide Number Placeholder 4"
            };

            var nonVisualShapeDrawingProperties5 = new NonVisualShapeDrawingProperties();
            var shapeLocks5 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties5.Append(shapeLocks5);

            var applicationNonVisualDrawingProperties7 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape5 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 12U
            };

            applicationNonVisualDrawingProperties7.Append(placeholderShape5);

            nonVisualShapeProperties5.Append(nonVisualDrawingProperties7);
            nonVisualShapeProperties5.Append(nonVisualShapeDrawingProperties5);
            nonVisualShapeProperties5.Append(applicationNonVisualDrawingProperties7);
            var shapeProperties5 = new ShapeProperties();

            var textBody5 = new TextBody();
            var bodyProperties5 = new A.BodyProperties();
            var listStyle5 = new A.ListStyle();

            var paragraph5 = new A.Paragraph();

            var field2 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties3 = new A.RunProperties() {Language = "en-GB"};
            runProperties3.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text3 = new A.Text();
            text3.Text = "‹#›";

            field2.Append(runProperties3);
            field2.Append(text3);
            var endParagraphRunProperties5 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph5.Append(field2);
            paragraph5.Append(endParagraphRunProperties5);

            textBody5.Append(bodyProperties5);
            textBody5.Append(listStyle5);
            textBody5.Append(paragraph5);

            shape5.Append(nonVisualShapeProperties5);
            shape5.Append(shapeProperties5);
            shape5.Append(textBody5);

            shapeTree2.Append(nonVisualGroupShapeProperties2);
            shapeTree2.Append(groupShapeProperties2);
            shapeTree2.Append(shape2);
            shapeTree2.Append(shape3);
            shapeTree2.Append(shape4);
            shapeTree2.Append(shape5);

            var commonSlideDataExtensionList2 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension2 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId2 = new P14.CreationId() {Val = (UInt32Value) 2908343129U};
            creationId2.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension2.Append(creationId2);

            commonSlideDataExtensionList2.Append(commonSlideDataExtension2);

            commonSlideData2.Append(shapeTree2);
            commonSlideData2.Append(commonSlideDataExtensionList2);

            var colorMapOverride2 = new ColorMapOverride();
            var masterColorMapping2 = new A.MasterColorMapping();

            colorMapOverride2.Append(masterColorMapping2);

            slideLayout1.Append(commonSlideData2);
            slideLayout1.Append(colorMapOverride2);

            slideLayoutPart1.SlideLayout = slideLayout1;
        }

        // Generates content of slideMasterPart1.
        private static void GenerateSlideMasterPart1Content(SlideMasterPart slideMasterPart1)
        {
            var slideMaster1 = new SlideMaster();
            slideMaster1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideMaster1.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideMaster1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData3 = new CommonSlideData();

            var background1 = new Background();

            var backgroundStyleReference1 = new BackgroundStyleReference() {Index = (UInt32Value) 1001U};
            var schemeColor10 = new A.SchemeColor() {Val = A.SchemeColorValues.Background1};

            backgroundStyleReference1.Append(schemeColor10);

            background1.Append(backgroundStyleReference1);

            var shapeTree3 = new ShapeTree();

            var nonVisualGroupShapeProperties3 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties8 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties3 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties8 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties3.Append(nonVisualDrawingProperties8);
            nonVisualGroupShapeProperties3.Append(nonVisualGroupShapeDrawingProperties3);
            nonVisualGroupShapeProperties3.Append(applicationNonVisualDrawingProperties8);

            var groupShapeProperties3 = new GroupShapeProperties();

            var transformGroup3 = new A.TransformGroup();
            var offset3 = new A.Offset() {X = 0L, Y = 0L};
            var extents3 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset3 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents3 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup3.Append(offset3);
            transformGroup3.Append(extents3);
            transformGroup3.Append(childOffset3);
            transformGroup3.Append(childExtents3);

            groupShapeProperties3.Append(transformGroup3);

            var shape6 = new Shape();

            var nonVisualShapeProperties6 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties9 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 2U,
                Name = "Title Placeholder 1"
            };

            var nonVisualShapeDrawingProperties6 = new NonVisualShapeDrawingProperties();
            var shapeLocks6 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties6.Append(shapeLocks6);

            var applicationNonVisualDrawingProperties9 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape6 = new PlaceholderShape() {Type = PlaceholderValues.Title};

            applicationNonVisualDrawingProperties9.Append(placeholderShape6);

            nonVisualShapeProperties6.Append(nonVisualDrawingProperties9);
            nonVisualShapeProperties6.Append(nonVisualShapeDrawingProperties6);
            nonVisualShapeProperties6.Append(applicationNonVisualDrawingProperties9);

            var shapeProperties6 = new ShapeProperties();

            var transform2D1 = new A.Transform2D();
            var offset4 = new A.Offset() {X = 457200L, Y = 274638L};
            var extents4 = new A.Extents() {Cx = 8229600L, Cy = 1143000L};

            transform2D1.Append(offset4);
            transform2D1.Append(extents4);

            var presetGeometry1 = new A.PresetGeometry() {Preset = A.ShapeTypeValues.Rectangle};
            var adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties6.Append(transform2D1);
            shapeProperties6.Append(presetGeometry1);

            var textBody6 = new TextBody();

            var bodyProperties6 = new A.BodyProperties() {
                Vertical = A.TextVerticalValues.Horizontal,
                LeftInset = 91440,
                TopInset = 45720,
                RightInset = 91440,
                BottomInset = 45720,
                RightToLeftColumns = false,
                Anchor = A.TextAnchoringTypeValues.Center
            };
            var normalAutoFit1 = new A.NormalAutoFit();

            bodyProperties6.Append(normalAutoFit1);
            var listStyle6 = new A.ListStyle();

            var paragraph6 = new A.Paragraph();

            var run2 = new A.Run();

            var runProperties4 = new A.RunProperties() {Language = "en-US"};
            runProperties4.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text4 = new A.Text();
            text4.Text = "Click to edit Master title style";

            run2.Append(runProperties4);
            run2.Append(text4);
            var endParagraphRunProperties6 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph6.Append(run2);
            paragraph6.Append(endParagraphRunProperties6);

            textBody6.Append(bodyProperties6);
            textBody6.Append(listStyle6);
            textBody6.Append(paragraph6);

            shape6.Append(nonVisualShapeProperties6);
            shape6.Append(shapeProperties6);
            shape6.Append(textBody6);

            var shape7 = new Shape();

            var nonVisualShapeProperties7 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties10 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Text Placeholder 2"
            };

            var nonVisualShapeDrawingProperties7 = new NonVisualShapeDrawingProperties();
            var shapeLocks7 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties7.Append(shapeLocks7);

            var applicationNonVisualDrawingProperties10 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape7 = new PlaceholderShape() {Type = PlaceholderValues.Body, Index = (UInt32Value) 1U};

            applicationNonVisualDrawingProperties10.Append(placeholderShape7);

            nonVisualShapeProperties7.Append(nonVisualDrawingProperties10);
            nonVisualShapeProperties7.Append(nonVisualShapeDrawingProperties7);
            nonVisualShapeProperties7.Append(applicationNonVisualDrawingProperties10);

            var shapeProperties7 = new ShapeProperties();

            var transform2D2 = new A.Transform2D();
            var offset5 = new A.Offset() {X = 457200L, Y = 1600200L};
            var extents5 = new A.Extents() {Cx = 8229600L, Cy = 4525963L};

            transform2D2.Append(offset5);
            transform2D2.Append(extents5);

            var presetGeometry2 = new A.PresetGeometry() {Preset = A.ShapeTypeValues.Rectangle};
            var adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);

            shapeProperties7.Append(transform2D2);
            shapeProperties7.Append(presetGeometry2);

            var textBody7 = new TextBody();

            var bodyProperties7 = new A.BodyProperties() {
                Vertical = A.TextVerticalValues.Horizontal,
                LeftInset = 91440,
                TopInset = 45720,
                RightInset = 91440,
                BottomInset = 45720,
                RightToLeftColumns = false
            };
            var normalAutoFit2 = new A.NormalAutoFit();

            bodyProperties7.Append(normalAutoFit2);
            var listStyle7 = new A.ListStyle();

            var paragraph7 = new A.Paragraph();
            var paragraphProperties1 = new A.ParagraphProperties() {Level = 0};

            var run3 = new A.Run();

            var runProperties5 = new A.RunProperties() {Language = "en-US"};
            runProperties5.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text5 = new A.Text();
            text5.Text = "Click to edit Master text styles";

            run3.Append(runProperties5);
            run3.Append(text5);

            paragraph7.Append(paragraphProperties1);
            paragraph7.Append(run3);

            var paragraph8 = new A.Paragraph();
            var paragraphProperties2 = new A.ParagraphProperties() {Level = 1};

            var run4 = new A.Run();

            var runProperties6 = new A.RunProperties() {Language = "en-US"};
            runProperties6.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text6 = new A.Text();
            text6.Text = "Second level";

            run4.Append(runProperties6);
            run4.Append(text6);

            paragraph8.Append(paragraphProperties2);
            paragraph8.Append(run4);

            var paragraph9 = new A.Paragraph();
            var paragraphProperties3 = new A.ParagraphProperties() {Level = 2};

            var run5 = new A.Run();

            var runProperties7 = new A.RunProperties() {Language = "en-US"};
            runProperties7.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text7 = new A.Text();
            text7.Text = "Third level";

            run5.Append(runProperties7);
            run5.Append(text7);

            paragraph9.Append(paragraphProperties3);
            paragraph9.Append(run5);

            var paragraph10 = new A.Paragraph();
            var paragraphProperties4 = new A.ParagraphProperties() {Level = 3};

            var run6 = new A.Run();

            var runProperties8 = new A.RunProperties() {Language = "en-US"};
            runProperties8.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text8 = new A.Text();
            text8.Text = "Fourth level";

            run6.Append(runProperties8);
            run6.Append(text8);

            paragraph10.Append(paragraphProperties4);
            paragraph10.Append(run6);

            var paragraph11 = new A.Paragraph();
            var paragraphProperties5 = new A.ParagraphProperties() {Level = 4};

            var run7 = new A.Run();

            var runProperties9 = new A.RunProperties() {Language = "en-US"};
            runProperties9.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text9 = new A.Text();
            text9.Text = "Fifth level";

            run7.Append(runProperties9);
            run7.Append(text9);
            var endParagraphRunProperties7 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph11.Append(paragraphProperties5);
            paragraph11.Append(run7);
            paragraph11.Append(endParagraphRunProperties7);

            textBody7.Append(bodyProperties7);
            textBody7.Append(listStyle7);
            textBody7.Append(paragraph7);
            textBody7.Append(paragraph8);
            textBody7.Append(paragraph9);
            textBody7.Append(paragraph10);
            textBody7.Append(paragraph11);

            shape7.Append(nonVisualShapeProperties7);
            shape7.Append(shapeProperties7);
            shape7.Append(textBody7);

            var shape8 = new Shape();

            var nonVisualShapeProperties8 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties11 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Date Placeholder 3"
            };

            var nonVisualShapeDrawingProperties8 = new NonVisualShapeDrawingProperties();
            var shapeLocks8 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties8.Append(shapeLocks8);

            var applicationNonVisualDrawingProperties11 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape8 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 2U
            };

            applicationNonVisualDrawingProperties11.Append(placeholderShape8);

            nonVisualShapeProperties8.Append(nonVisualDrawingProperties11);
            nonVisualShapeProperties8.Append(nonVisualShapeDrawingProperties8);
            nonVisualShapeProperties8.Append(applicationNonVisualDrawingProperties11);

            var shapeProperties8 = new ShapeProperties();

            var transform2D3 = new A.Transform2D();
            var offset6 = new A.Offset() {X = 457200L, Y = 6356350L};
            var extents6 = new A.Extents() {Cx = 2133600L, Cy = 365125L};

            transform2D3.Append(offset6);
            transform2D3.Append(extents6);

            var presetGeometry3 = new A.PresetGeometry() {Preset = A.ShapeTypeValues.Rectangle};
            var adjustValueList3 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList3);

            shapeProperties8.Append(transform2D3);
            shapeProperties8.Append(presetGeometry3);

            var textBody8 = new TextBody();
            var bodyProperties8 = new A.BodyProperties() {
                Vertical = A.TextVerticalValues.Horizontal,
                LeftInset = 91440,
                TopInset = 45720,
                RightInset = 91440,
                BottomInset = 45720,
                RightToLeftColumns = false,
                Anchor = A.TextAnchoringTypeValues.Center
            };

            var listStyle8 = new A.ListStyle();

            var level1ParagraphProperties2 = new A.Level1ParagraphProperties() {
                Alignment = A.TextAlignmentTypeValues.Left
            };

            var defaultRunProperties11 = new A.DefaultRunProperties() {FontSize = 1200};

            var solidFill10 = new A.SolidFill();

            var schemeColor11 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint1 = new A.Tint() {Val = 75000};

            schemeColor11.Append(tint1);

            solidFill10.Append(schemeColor11);

            defaultRunProperties11.Append(solidFill10);

            level1ParagraphProperties2.Append(defaultRunProperties11);

            listStyle8.Append(level1ParagraphProperties2);

            var paragraph12 = new A.Paragraph();

            var field3 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties10 = new A.RunProperties() {Language = "en-GB"};
            runProperties10.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text10 = new A.Text();
            text10.Text = "04/06/2014";

            field3.Append(runProperties10);
            field3.Append(text10);
            var endParagraphRunProperties8 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph12.Append(field3);
            paragraph12.Append(endParagraphRunProperties8);

            textBody8.Append(bodyProperties8);
            textBody8.Append(listStyle8);
            textBody8.Append(paragraph12);

            shape8.Append(nonVisualShapeProperties8);
            shape8.Append(shapeProperties8);
            shape8.Append(textBody8);

            var shape9 = new Shape();

            var nonVisualShapeProperties9 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties12 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 5U,
                Name = "Footer Placeholder 4"
            };

            var nonVisualShapeDrawingProperties9 = new NonVisualShapeDrawingProperties();
            var shapeLocks9 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties9.Append(shapeLocks9);

            var applicationNonVisualDrawingProperties12 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape9 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 3U
            };

            applicationNonVisualDrawingProperties12.Append(placeholderShape9);

            nonVisualShapeProperties9.Append(nonVisualDrawingProperties12);
            nonVisualShapeProperties9.Append(nonVisualShapeDrawingProperties9);
            nonVisualShapeProperties9.Append(applicationNonVisualDrawingProperties12);

            var shapeProperties9 = new ShapeProperties();

            var transform2D4 = new A.Transform2D();
            var offset7 = new A.Offset() {X = 3124200L, Y = 6356350L};
            var extents7 = new A.Extents() {Cx = 2895600L, Cy = 365125L};

            transform2D4.Append(offset7);
            transform2D4.Append(extents7);

            var presetGeometry4 = new A.PresetGeometry() {Preset = A.ShapeTypeValues.Rectangle};
            var adjustValueList4 = new A.AdjustValueList();

            presetGeometry4.Append(adjustValueList4);

            shapeProperties9.Append(transform2D4);
            shapeProperties9.Append(presetGeometry4);

            var textBody9 = new TextBody();
            var bodyProperties9 = new A.BodyProperties() {
                Vertical = A.TextVerticalValues.Horizontal,
                LeftInset = 91440,
                TopInset = 45720,
                RightInset = 91440,
                BottomInset = 45720,
                RightToLeftColumns = false,
                Anchor = A.TextAnchoringTypeValues.Center
            };

            var listStyle9 = new A.ListStyle();

            var level1ParagraphProperties3 = new A.Level1ParagraphProperties() {
                Alignment = A.TextAlignmentTypeValues.Center
            };

            var defaultRunProperties12 = new A.DefaultRunProperties() {FontSize = 1200};

            var solidFill11 = new A.SolidFill();

            var schemeColor12 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint2 = new A.Tint() {Val = 75000};

            schemeColor12.Append(tint2);

            solidFill11.Append(schemeColor12);

            defaultRunProperties12.Append(solidFill11);

            level1ParagraphProperties3.Append(defaultRunProperties12);

            listStyle9.Append(level1ParagraphProperties3);

            var paragraph13 = new A.Paragraph();
            var endParagraphRunProperties9 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph13.Append(endParagraphRunProperties9);

            textBody9.Append(bodyProperties9);
            textBody9.Append(listStyle9);
            textBody9.Append(paragraph13);

            shape9.Append(nonVisualShapeProperties9);
            shape9.Append(shapeProperties9);
            shape9.Append(textBody9);

            var shape10 = new Shape();

            var nonVisualShapeProperties10 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties13 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 6U,
                Name = "Slide Number Placeholder 5"
            };

            var nonVisualShapeDrawingProperties10 = new NonVisualShapeDrawingProperties();
            var shapeLocks10 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties10.Append(shapeLocks10);

            var applicationNonVisualDrawingProperties13 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape10 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 4U
            };

            applicationNonVisualDrawingProperties13.Append(placeholderShape10);

            nonVisualShapeProperties10.Append(nonVisualDrawingProperties13);
            nonVisualShapeProperties10.Append(nonVisualShapeDrawingProperties10);
            nonVisualShapeProperties10.Append(applicationNonVisualDrawingProperties13);

            var shapeProperties10 = new ShapeProperties();

            var transform2D5 = new A.Transform2D();
            var offset8 = new A.Offset() {X = 6553200L, Y = 6356350L};
            var extents8 = new A.Extents() {Cx = 2133600L, Cy = 365125L};

            transform2D5.Append(offset8);
            transform2D5.Append(extents8);

            var presetGeometry5 = new A.PresetGeometry() {Preset = A.ShapeTypeValues.Rectangle};
            var adjustValueList5 = new A.AdjustValueList();

            presetGeometry5.Append(adjustValueList5);

            shapeProperties10.Append(transform2D5);
            shapeProperties10.Append(presetGeometry5);

            var textBody10 = new TextBody();
            var bodyProperties10 = new A.BodyProperties() {
                Vertical = A.TextVerticalValues.Horizontal,
                LeftInset = 91440,
                TopInset = 45720,
                RightInset = 91440,
                BottomInset = 45720,
                RightToLeftColumns = false,
                Anchor = A.TextAnchoringTypeValues.Center
            };

            var listStyle10 = new A.ListStyle();

            var level1ParagraphProperties4 = new A.Level1ParagraphProperties() {
                Alignment = A.TextAlignmentTypeValues.Right
            };

            var defaultRunProperties13 = new A.DefaultRunProperties() {FontSize = 1200};

            var solidFill12 = new A.SolidFill();

            var schemeColor13 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint3 = new A.Tint() {Val = 75000};

            schemeColor13.Append(tint3);

            solidFill12.Append(schemeColor13);

            defaultRunProperties13.Append(solidFill12);

            level1ParagraphProperties4.Append(defaultRunProperties13);

            listStyle10.Append(level1ParagraphProperties4);

            var paragraph14 = new A.Paragraph();

            var field4 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties11 = new A.RunProperties() {Language = "en-GB"};
            runProperties11.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text11 = new A.Text();
            text11.Text = "‹#›";

            field4.Append(runProperties11);
            field4.Append(text11);
            var endParagraphRunProperties10 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph14.Append(field4);
            paragraph14.Append(endParagraphRunProperties10);

            textBody10.Append(bodyProperties10);
            textBody10.Append(listStyle10);
            textBody10.Append(paragraph14);

            shape10.Append(nonVisualShapeProperties10);
            shape10.Append(shapeProperties10);
            shape10.Append(textBody10);

            shapeTree3.Append(nonVisualGroupShapeProperties3);
            shapeTree3.Append(groupShapeProperties3);
            shapeTree3.Append(shape6);
            shapeTree3.Append(shape7);
            shapeTree3.Append(shape8);
            shapeTree3.Append(shape9);
            shapeTree3.Append(shape10);

            var commonSlideDataExtensionList3 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension3 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId3 = new P14.CreationId() {Val = (UInt32Value) 616110946U};
            creationId3.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension3.Append(creationId3);

            commonSlideDataExtensionList3.Append(commonSlideDataExtension3);

            commonSlideData3.Append(background1);
            commonSlideData3.Append(shapeTree3);
            commonSlideData3.Append(commonSlideDataExtensionList3);
            var colorMap1 = new ColorMap() {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            };

            var slideLayoutIdList1 = new SlideLayoutIdList();
            var slideLayoutId1 = new SlideLayoutId() {Id = (UInt32Value) 2147483649U, RelationshipId = "rId1"};
            var slideLayoutId2 = new SlideLayoutId() {Id = (UInt32Value) 2147483650U, RelationshipId = "rId2"};
            var slideLayoutId3 = new SlideLayoutId() {Id = (UInt32Value) 2147483651U, RelationshipId = "rId3"};
            var slideLayoutId4 = new SlideLayoutId() {Id = (UInt32Value) 2147483652U, RelationshipId = "rId4"};
            var slideLayoutId5 = new SlideLayoutId() {Id = (UInt32Value) 2147483653U, RelationshipId = "rId5"};
            var slideLayoutId6 = new SlideLayoutId() {Id = (UInt32Value) 2147483654U, RelationshipId = "rId6"};
            var slideLayoutId7 = new SlideLayoutId() {Id = (UInt32Value) 2147483655U, RelationshipId = "rId7"};
            var slideLayoutId8 = new SlideLayoutId() {Id = (UInt32Value) 2147483656U, RelationshipId = "rId8"};
            var slideLayoutId9 = new SlideLayoutId() {Id = (UInt32Value) 2147483657U, RelationshipId = "rId9"};
            var slideLayoutId10 = new SlideLayoutId() {Id = (UInt32Value) 2147483658U, RelationshipId = "rId10"};
            var slideLayoutId11 = new SlideLayoutId() {Id = (UInt32Value) 2147483659U, RelationshipId = "rId11"};

            slideLayoutIdList1.Append(slideLayoutId1);
            slideLayoutIdList1.Append(slideLayoutId2);
            slideLayoutIdList1.Append(slideLayoutId3);
            slideLayoutIdList1.Append(slideLayoutId4);
            slideLayoutIdList1.Append(slideLayoutId5);
            slideLayoutIdList1.Append(slideLayoutId6);
            slideLayoutIdList1.Append(slideLayoutId7);
            slideLayoutIdList1.Append(slideLayoutId8);
            slideLayoutIdList1.Append(slideLayoutId9);
            slideLayoutIdList1.Append(slideLayoutId10);
            slideLayoutIdList1.Append(slideLayoutId11);

            var textStyles1 = new TextStyles();

            var titleStyle1 = new TitleStyle();

            var level1ParagraphProperties5 = new A.Level1ParagraphProperties() {
                Alignment = A.TextAlignmentTypeValues.Center,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var spaceBefore1 = new A.SpaceBefore();
            var spacingPercent1 = new A.SpacingPercent() {Val = 0};

            spaceBefore1.Append(spacingPercent1);
            var noBullet1 = new A.NoBullet();

            var defaultRunProperties14 = new A.DefaultRunProperties() {FontSize = 4400, Kerning = 1200};

            var solidFill13 = new A.SolidFill();
            var schemeColor14 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill13.Append(schemeColor14);
            var latinFont10 = new A.LatinFont() {Typeface = "+mj-lt"};
            var eastAsianFont10 = new A.EastAsianFont() {Typeface = "+mj-ea"};
            var complexScriptFont10 = new A.ComplexScriptFont() {Typeface = "+mj-cs"};

            defaultRunProperties14.Append(solidFill13);
            defaultRunProperties14.Append(latinFont10);
            defaultRunProperties14.Append(eastAsianFont10);
            defaultRunProperties14.Append(complexScriptFont10);

            level1ParagraphProperties5.Append(spaceBefore1);
            level1ParagraphProperties5.Append(noBullet1);
            level1ParagraphProperties5.Append(defaultRunProperties14);

            titleStyle1.Append(level1ParagraphProperties5);

            var bodyStyle1 = new BodyStyle();

            var level1ParagraphProperties6 = new A.Level1ParagraphProperties() {
                LeftMargin = 342900,
                Indent = -342900,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var spaceBefore2 = new A.SpaceBefore();
            var spacingPercent2 = new A.SpacingPercent() {Val = 20000};

            spaceBefore2.Append(spacingPercent2);
            var bulletFont1 = new A.BulletFont() {
                Typeface = "Arial",
                Panose = "020B0604020202020204",
                PitchFamily = 34,
                CharacterSet = 0
            };
            var characterBullet1 = new A.CharacterBullet() {Char = "•"};

            var defaultRunProperties15 = new A.DefaultRunProperties() {FontSize = 3200, Kerning = 1200};

            var solidFill14 = new A.SolidFill();
            var schemeColor15 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill14.Append(schemeColor15);
            var latinFont11 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont11 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont11 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties15.Append(solidFill14);
            defaultRunProperties15.Append(latinFont11);
            defaultRunProperties15.Append(eastAsianFont11);
            defaultRunProperties15.Append(complexScriptFont11);

            level1ParagraphProperties6.Append(spaceBefore2);
            level1ParagraphProperties6.Append(bulletFont1);
            level1ParagraphProperties6.Append(characterBullet1);
            level1ParagraphProperties6.Append(defaultRunProperties15);

            var level2ParagraphProperties2 = new A.Level2ParagraphProperties() {
                LeftMargin = 742950,
                Indent = -285750,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var spaceBefore3 = new A.SpaceBefore();
            var spacingPercent3 = new A.SpacingPercent() {Val = 20000};

            spaceBefore3.Append(spacingPercent3);
            var bulletFont2 = new A.BulletFont() {
                Typeface = "Arial",
                Panose = "020B0604020202020204",
                PitchFamily = 34,
                CharacterSet = 0
            };
            var characterBullet2 = new A.CharacterBullet() {Char = "–"};

            var defaultRunProperties16 = new A.DefaultRunProperties() {FontSize = 2800, Kerning = 1200};

            var solidFill15 = new A.SolidFill();
            var schemeColor16 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill15.Append(schemeColor16);
            var latinFont12 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont12 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont12 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties16.Append(solidFill15);
            defaultRunProperties16.Append(latinFont12);
            defaultRunProperties16.Append(eastAsianFont12);
            defaultRunProperties16.Append(complexScriptFont12);

            level2ParagraphProperties2.Append(spaceBefore3);
            level2ParagraphProperties2.Append(bulletFont2);
            level2ParagraphProperties2.Append(characterBullet2);
            level2ParagraphProperties2.Append(defaultRunProperties16);

            var level3ParagraphProperties2 = new A.Level3ParagraphProperties() {
                LeftMargin = 1143000,
                Indent = -228600,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var spaceBefore4 = new A.SpaceBefore();
            var spacingPercent4 = new A.SpacingPercent() {Val = 20000};

            spaceBefore4.Append(spacingPercent4);
            var bulletFont3 = new A.BulletFont() {
                Typeface = "Arial",
                Panose = "020B0604020202020204",
                PitchFamily = 34,
                CharacterSet = 0
            };
            var characterBullet3 = new A.CharacterBullet() {Char = "•"};

            var defaultRunProperties17 = new A.DefaultRunProperties() {FontSize = 2400, Kerning = 1200};

            var solidFill16 = new A.SolidFill();
            var schemeColor17 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill16.Append(schemeColor17);
            var latinFont13 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont13 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont13 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties17.Append(solidFill16);
            defaultRunProperties17.Append(latinFont13);
            defaultRunProperties17.Append(eastAsianFont13);
            defaultRunProperties17.Append(complexScriptFont13);

            level3ParagraphProperties2.Append(spaceBefore4);
            level3ParagraphProperties2.Append(bulletFont3);
            level3ParagraphProperties2.Append(characterBullet3);
            level3ParagraphProperties2.Append(defaultRunProperties17);

            var level4ParagraphProperties2 = new A.Level4ParagraphProperties() {
                LeftMargin = 1600200,
                Indent = -228600,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var spaceBefore5 = new A.SpaceBefore();
            var spacingPercent5 = new A.SpacingPercent() {Val = 20000};

            spaceBefore5.Append(spacingPercent5);
            var bulletFont4 = new A.BulletFont() {
                Typeface = "Arial",
                Panose = "020B0604020202020204",
                PitchFamily = 34,
                CharacterSet = 0
            };
            var characterBullet4 = new A.CharacterBullet() {Char = "–"};

            var defaultRunProperties18 = new A.DefaultRunProperties() {FontSize = 2000, Kerning = 1200};

            var solidFill17 = new A.SolidFill();
            var schemeColor18 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill17.Append(schemeColor18);
            var latinFont14 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont14 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont14 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties18.Append(solidFill17);
            defaultRunProperties18.Append(latinFont14);
            defaultRunProperties18.Append(eastAsianFont14);
            defaultRunProperties18.Append(complexScriptFont14);

            level4ParagraphProperties2.Append(spaceBefore5);
            level4ParagraphProperties2.Append(bulletFont4);
            level4ParagraphProperties2.Append(characterBullet4);
            level4ParagraphProperties2.Append(defaultRunProperties18);

            var level5ParagraphProperties2 = new A.Level5ParagraphProperties() {
                LeftMargin = 2057400,
                Indent = -228600,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var spaceBefore6 = new A.SpaceBefore();
            var spacingPercent6 = new A.SpacingPercent() {Val = 20000};

            spaceBefore6.Append(spacingPercent6);
            var bulletFont5 = new A.BulletFont() {
                Typeface = "Arial",
                Panose = "020B0604020202020204",
                PitchFamily = 34,
                CharacterSet = 0
            };
            var characterBullet5 = new A.CharacterBullet() {Char = "»"};

            var defaultRunProperties19 = new A.DefaultRunProperties() {FontSize = 2000, Kerning = 1200};

            var solidFill18 = new A.SolidFill();
            var schemeColor19 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill18.Append(schemeColor19);
            var latinFont15 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont15 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont15 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties19.Append(solidFill18);
            defaultRunProperties19.Append(latinFont15);
            defaultRunProperties19.Append(eastAsianFont15);
            defaultRunProperties19.Append(complexScriptFont15);

            level5ParagraphProperties2.Append(spaceBefore6);
            level5ParagraphProperties2.Append(bulletFont5);
            level5ParagraphProperties2.Append(characterBullet5);
            level5ParagraphProperties2.Append(defaultRunProperties19);

            var level6ParagraphProperties2 = new A.Level6ParagraphProperties() {
                LeftMargin = 2514600,
                Indent = -228600,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var spaceBefore7 = new A.SpaceBefore();
            var spacingPercent7 = new A.SpacingPercent() {Val = 20000};

            spaceBefore7.Append(spacingPercent7);
            var bulletFont6 = new A.BulletFont() {
                Typeface = "Arial",
                Panose = "020B0604020202020204",
                PitchFamily = 34,
                CharacterSet = 0
            };
            var characterBullet6 = new A.CharacterBullet() {Char = "•"};

            var defaultRunProperties20 = new A.DefaultRunProperties() {FontSize = 2000, Kerning = 1200};

            var solidFill19 = new A.SolidFill();
            var schemeColor20 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill19.Append(schemeColor20);
            var latinFont16 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont16 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont16 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties20.Append(solidFill19);
            defaultRunProperties20.Append(latinFont16);
            defaultRunProperties20.Append(eastAsianFont16);
            defaultRunProperties20.Append(complexScriptFont16);

            level6ParagraphProperties2.Append(spaceBefore7);
            level6ParagraphProperties2.Append(bulletFont6);
            level6ParagraphProperties2.Append(characterBullet6);
            level6ParagraphProperties2.Append(defaultRunProperties20);

            var level7ParagraphProperties2 = new A.Level7ParagraphProperties() {
                LeftMargin = 2971800,
                Indent = -228600,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var spaceBefore8 = new A.SpaceBefore();
            var spacingPercent8 = new A.SpacingPercent() {Val = 20000};

            spaceBefore8.Append(spacingPercent8);
            var bulletFont7 = new A.BulletFont() {
                Typeface = "Arial",
                Panose = "020B0604020202020204",
                PitchFamily = 34,
                CharacterSet = 0
            };
            var characterBullet7 = new A.CharacterBullet() {Char = "•"};

            var defaultRunProperties21 = new A.DefaultRunProperties() {FontSize = 2000, Kerning = 1200};

            var solidFill20 = new A.SolidFill();
            var schemeColor21 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill20.Append(schemeColor21);
            var latinFont17 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont17 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont17 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties21.Append(solidFill20);
            defaultRunProperties21.Append(latinFont17);
            defaultRunProperties21.Append(eastAsianFont17);
            defaultRunProperties21.Append(complexScriptFont17);

            level7ParagraphProperties2.Append(spaceBefore8);
            level7ParagraphProperties2.Append(bulletFont7);
            level7ParagraphProperties2.Append(characterBullet7);
            level7ParagraphProperties2.Append(defaultRunProperties21);

            var level8ParagraphProperties2 = new A.Level8ParagraphProperties() {
                LeftMargin = 3429000,
                Indent = -228600,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var spaceBefore9 = new A.SpaceBefore();
            var spacingPercent9 = new A.SpacingPercent() {Val = 20000};

            spaceBefore9.Append(spacingPercent9);
            var bulletFont8 = new A.BulletFont() {
                Typeface = "Arial",
                Panose = "020B0604020202020204",
                PitchFamily = 34,
                CharacterSet = 0
            };
            var characterBullet8 = new A.CharacterBullet() {Char = "•"};

            var defaultRunProperties22 = new A.DefaultRunProperties() {FontSize = 2000, Kerning = 1200};

            var solidFill21 = new A.SolidFill();
            var schemeColor22 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill21.Append(schemeColor22);
            var latinFont18 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont18 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont18 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties22.Append(solidFill21);
            defaultRunProperties22.Append(latinFont18);
            defaultRunProperties22.Append(eastAsianFont18);
            defaultRunProperties22.Append(complexScriptFont18);

            level8ParagraphProperties2.Append(spaceBefore9);
            level8ParagraphProperties2.Append(bulletFont8);
            level8ParagraphProperties2.Append(characterBullet8);
            level8ParagraphProperties2.Append(defaultRunProperties22);

            var level9ParagraphProperties2 = new A.Level9ParagraphProperties() {
                LeftMargin = 3886200,
                Indent = -228600,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var spaceBefore10 = new A.SpaceBefore();
            var spacingPercent10 = new A.SpacingPercent() {Val = 20000};

            spaceBefore10.Append(spacingPercent10);
            var bulletFont9 = new A.BulletFont() {
                Typeface = "Arial",
                Panose = "020B0604020202020204",
                PitchFamily = 34,
                CharacterSet = 0
            };
            var characterBullet9 = new A.CharacterBullet() {Char = "•"};

            var defaultRunProperties23 = new A.DefaultRunProperties() {FontSize = 2000, Kerning = 1200};

            var solidFill22 = new A.SolidFill();
            var schemeColor23 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill22.Append(schemeColor23);
            var latinFont19 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont19 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont19 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties23.Append(solidFill22);
            defaultRunProperties23.Append(latinFont19);
            defaultRunProperties23.Append(eastAsianFont19);
            defaultRunProperties23.Append(complexScriptFont19);

            level9ParagraphProperties2.Append(spaceBefore10);
            level9ParagraphProperties2.Append(bulletFont9);
            level9ParagraphProperties2.Append(characterBullet9);
            level9ParagraphProperties2.Append(defaultRunProperties23);

            bodyStyle1.Append(level1ParagraphProperties6);
            bodyStyle1.Append(level2ParagraphProperties2);
            bodyStyle1.Append(level3ParagraphProperties2);
            bodyStyle1.Append(level4ParagraphProperties2);
            bodyStyle1.Append(level5ParagraphProperties2);
            bodyStyle1.Append(level6ParagraphProperties2);
            bodyStyle1.Append(level7ParagraphProperties2);
            bodyStyle1.Append(level8ParagraphProperties2);
            bodyStyle1.Append(level9ParagraphProperties2);

            var otherStyle1 = new OtherStyle();

            var defaultParagraphProperties2 = new A.DefaultParagraphProperties();
            var defaultRunProperties24 = new A.DefaultRunProperties() {Language = "nl-NL"};

            defaultParagraphProperties2.Append(defaultRunProperties24);

            var level1ParagraphProperties7 = new A.Level1ParagraphProperties() {
                LeftMargin = 0,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties25 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill23 = new A.SolidFill();
            var schemeColor24 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill23.Append(schemeColor24);
            var latinFont20 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont20 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont20 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties25.Append(solidFill23);
            defaultRunProperties25.Append(latinFont20);
            defaultRunProperties25.Append(eastAsianFont20);
            defaultRunProperties25.Append(complexScriptFont20);

            level1ParagraphProperties7.Append(defaultRunProperties25);

            var level2ParagraphProperties3 = new A.Level2ParagraphProperties() {
                LeftMargin = 457200,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties26 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill24 = new A.SolidFill();
            var schemeColor25 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill24.Append(schemeColor25);
            var latinFont21 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont21 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont21 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties26.Append(solidFill24);
            defaultRunProperties26.Append(latinFont21);
            defaultRunProperties26.Append(eastAsianFont21);
            defaultRunProperties26.Append(complexScriptFont21);

            level2ParagraphProperties3.Append(defaultRunProperties26);

            var level3ParagraphProperties3 = new A.Level3ParagraphProperties() {
                LeftMargin = 914400,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties27 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill25 = new A.SolidFill();
            var schemeColor26 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill25.Append(schemeColor26);
            var latinFont22 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont22 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont22 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties27.Append(solidFill25);
            defaultRunProperties27.Append(latinFont22);
            defaultRunProperties27.Append(eastAsianFont22);
            defaultRunProperties27.Append(complexScriptFont22);

            level3ParagraphProperties3.Append(defaultRunProperties27);

            var level4ParagraphProperties3 = new A.Level4ParagraphProperties() {
                LeftMargin = 1371600,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties28 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill26 = new A.SolidFill();
            var schemeColor27 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill26.Append(schemeColor27);
            var latinFont23 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont23 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont23 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties28.Append(solidFill26);
            defaultRunProperties28.Append(latinFont23);
            defaultRunProperties28.Append(eastAsianFont23);
            defaultRunProperties28.Append(complexScriptFont23);

            level4ParagraphProperties3.Append(defaultRunProperties28);

            var level5ParagraphProperties3 = new A.Level5ParagraphProperties() {
                LeftMargin = 1828800,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties29 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill27 = new A.SolidFill();
            var schemeColor28 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill27.Append(schemeColor28);
            var latinFont24 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont24 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont24 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties29.Append(solidFill27);
            defaultRunProperties29.Append(latinFont24);
            defaultRunProperties29.Append(eastAsianFont24);
            defaultRunProperties29.Append(complexScriptFont24);

            level5ParagraphProperties3.Append(defaultRunProperties29);

            var level6ParagraphProperties3 = new A.Level6ParagraphProperties() {
                LeftMargin = 2286000,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties30 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill28 = new A.SolidFill();
            var schemeColor29 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill28.Append(schemeColor29);
            var latinFont25 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont25 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont25 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties30.Append(solidFill28);
            defaultRunProperties30.Append(latinFont25);
            defaultRunProperties30.Append(eastAsianFont25);
            defaultRunProperties30.Append(complexScriptFont25);

            level6ParagraphProperties3.Append(defaultRunProperties30);

            var level7ParagraphProperties3 = new A.Level7ParagraphProperties() {
                LeftMargin = 2743200,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties31 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill29 = new A.SolidFill();
            var schemeColor30 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill29.Append(schemeColor30);
            var latinFont26 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont26 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont26 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties31.Append(solidFill29);
            defaultRunProperties31.Append(latinFont26);
            defaultRunProperties31.Append(eastAsianFont26);
            defaultRunProperties31.Append(complexScriptFont26);

            level7ParagraphProperties3.Append(defaultRunProperties31);

            var level8ParagraphProperties3 = new A.Level8ParagraphProperties() {
                LeftMargin = 3200400,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties32 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill30 = new A.SolidFill();
            var schemeColor31 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill30.Append(schemeColor31);
            var latinFont27 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont27 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont27 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties32.Append(solidFill30);
            defaultRunProperties32.Append(latinFont27);
            defaultRunProperties32.Append(eastAsianFont27);
            defaultRunProperties32.Append(complexScriptFont27);

            level8ParagraphProperties3.Append(defaultRunProperties32);

            var level9ParagraphProperties3 = new A.Level9ParagraphProperties() {
                LeftMargin = 3657600,
                Alignment = A.TextAlignmentTypeValues.Left,
                DefaultTabSize = 914400,
                RightToLeft = false,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };

            var defaultRunProperties33 = new A.DefaultRunProperties() {FontSize = 1800, Kerning = 1200};

            var solidFill31 = new A.SolidFill();
            var schemeColor32 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};

            solidFill31.Append(schemeColor32);
            var latinFont28 = new A.LatinFont() {Typeface = "+mn-lt"};
            var eastAsianFont28 = new A.EastAsianFont() {Typeface = "+mn-ea"};
            var complexScriptFont28 = new A.ComplexScriptFont() {Typeface = "+mn-cs"};

            defaultRunProperties33.Append(solidFill31);
            defaultRunProperties33.Append(latinFont28);
            defaultRunProperties33.Append(eastAsianFont28);
            defaultRunProperties33.Append(complexScriptFont28);

            level9ParagraphProperties3.Append(defaultRunProperties33);

            otherStyle1.Append(defaultParagraphProperties2);
            otherStyle1.Append(level1ParagraphProperties7);
            otherStyle1.Append(level2ParagraphProperties3);
            otherStyle1.Append(level3ParagraphProperties3);
            otherStyle1.Append(level4ParagraphProperties3);
            otherStyle1.Append(level5ParagraphProperties3);
            otherStyle1.Append(level6ParagraphProperties3);
            otherStyle1.Append(level7ParagraphProperties3);
            otherStyle1.Append(level8ParagraphProperties3);
            otherStyle1.Append(level9ParagraphProperties3);

            textStyles1.Append(titleStyle1);
            textStyles1.Append(bodyStyle1);
            textStyles1.Append(otherStyle1);

            slideMaster1.Append(commonSlideData3);
            slideMaster1.Append(colorMap1);
            slideMaster1.Append(slideLayoutIdList1);
            slideMaster1.Append(textStyles1);

            slideMasterPart1.SlideMaster = slideMaster1;
        }

        // Generates content of slideLayoutPart2.
        private static void GenerateSlideLayoutPart2Content(SlideLayoutPart slideLayoutPart2)
        {
            var slideLayout2 = new SlideLayout() {Type = SlideLayoutValues.ObjectText, Preserve = true};
            slideLayout2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout2.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout2.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData4 = new CommonSlideData() {Name = "Content with Caption"};

            var shapeTree4 = new ShapeTree();

            var nonVisualGroupShapeProperties4 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties14 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties4 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties14 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties4.Append(nonVisualDrawingProperties14);
            nonVisualGroupShapeProperties4.Append(nonVisualGroupShapeDrawingProperties4);
            nonVisualGroupShapeProperties4.Append(applicationNonVisualDrawingProperties14);

            var groupShapeProperties4 = new GroupShapeProperties();

            var transformGroup4 = new A.TransformGroup();
            var offset9 = new A.Offset() {X = 0L, Y = 0L};
            var extents9 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset4 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents4 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup4.Append(offset9);
            transformGroup4.Append(extents9);
            transformGroup4.Append(childOffset4);
            transformGroup4.Append(childExtents4);

            groupShapeProperties4.Append(transformGroup4);

            var shape11 = new Shape();

            var nonVisualShapeProperties11 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties15 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 2U,
                Name = "Title 1"
            };

            var nonVisualShapeDrawingProperties11 = new NonVisualShapeDrawingProperties();
            var shapeLocks11 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties11.Append(shapeLocks11);

            var applicationNonVisualDrawingProperties15 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape11 = new PlaceholderShape() {Type = PlaceholderValues.Title};

            applicationNonVisualDrawingProperties15.Append(placeholderShape11);

            nonVisualShapeProperties11.Append(nonVisualDrawingProperties15);
            nonVisualShapeProperties11.Append(nonVisualShapeDrawingProperties11);
            nonVisualShapeProperties11.Append(applicationNonVisualDrawingProperties15);

            var shapeProperties11 = new ShapeProperties();

            var transform2D6 = new A.Transform2D();
            var offset10 = new A.Offset() {X = 457200L, Y = 273050L};
            var extents10 = new A.Extents() {Cx = 3008313L, Cy = 1162050L};

            transform2D6.Append(offset10);
            transform2D6.Append(extents10);

            shapeProperties11.Append(transform2D6);

            var textBody11 = new TextBody();
            var bodyProperties11 = new A.BodyProperties() {Anchor = A.TextAnchoringTypeValues.Bottom};

            var listStyle11 = new A.ListStyle();

            var level1ParagraphProperties8 = new A.Level1ParagraphProperties() {
                Alignment = A.TextAlignmentTypeValues.Left
            };
            var defaultRunProperties34 = new A.DefaultRunProperties() {FontSize = 2000, Bold = true};

            level1ParagraphProperties8.Append(defaultRunProperties34);

            listStyle11.Append(level1ParagraphProperties8);

            var paragraph15 = new A.Paragraph();

            var run8 = new A.Run();

            var runProperties12 = new A.RunProperties() {Language = "en-US"};
            runProperties12.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text12 = new A.Text();
            text12.Text = "Click to edit Master title style";

            run8.Append(runProperties12);
            run8.Append(text12);
            var endParagraphRunProperties11 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph15.Append(run8);
            paragraph15.Append(endParagraphRunProperties11);

            textBody11.Append(bodyProperties11);
            textBody11.Append(listStyle11);
            textBody11.Append(paragraph15);

            shape11.Append(nonVisualShapeProperties11);
            shape11.Append(shapeProperties11);
            shape11.Append(textBody11);

            var shape12 = new Shape();

            var nonVisualShapeProperties12 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties16 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Content Placeholder 2"
            };

            var nonVisualShapeDrawingProperties12 = new NonVisualShapeDrawingProperties();
            var shapeLocks12 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties12.Append(shapeLocks12);

            var applicationNonVisualDrawingProperties16 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape12 = new PlaceholderShape() {Index = (UInt32Value) 1U};

            applicationNonVisualDrawingProperties16.Append(placeholderShape12);

            nonVisualShapeProperties12.Append(nonVisualDrawingProperties16);
            nonVisualShapeProperties12.Append(nonVisualShapeDrawingProperties12);
            nonVisualShapeProperties12.Append(applicationNonVisualDrawingProperties16);

            var shapeProperties12 = new ShapeProperties();

            var transform2D7 = new A.Transform2D();
            var offset11 = new A.Offset() {X = 3575050L, Y = 273050L};
            var extents11 = new A.Extents() {Cx = 5111750L, Cy = 5853113L};

            transform2D7.Append(offset11);
            transform2D7.Append(extents11);

            shapeProperties12.Append(transform2D7);

            var textBody12 = new TextBody();
            var bodyProperties12 = new A.BodyProperties();

            var listStyle12 = new A.ListStyle();

            var level1ParagraphProperties9 = new A.Level1ParagraphProperties();
            var defaultRunProperties35 = new A.DefaultRunProperties() {FontSize = 3200};

            level1ParagraphProperties9.Append(defaultRunProperties35);

            var level2ParagraphProperties4 = new A.Level2ParagraphProperties();
            var defaultRunProperties36 = new A.DefaultRunProperties() {FontSize = 2800};

            level2ParagraphProperties4.Append(defaultRunProperties36);

            var level3ParagraphProperties4 = new A.Level3ParagraphProperties();
            var defaultRunProperties37 = new A.DefaultRunProperties() {FontSize = 2400};

            level3ParagraphProperties4.Append(defaultRunProperties37);

            var level4ParagraphProperties4 = new A.Level4ParagraphProperties();
            var defaultRunProperties38 = new A.DefaultRunProperties() {FontSize = 2000};

            level4ParagraphProperties4.Append(defaultRunProperties38);

            var level5ParagraphProperties4 = new A.Level5ParagraphProperties();
            var defaultRunProperties39 = new A.DefaultRunProperties() {FontSize = 2000};

            level5ParagraphProperties4.Append(defaultRunProperties39);

            var level6ParagraphProperties4 = new A.Level6ParagraphProperties();
            var defaultRunProperties40 = new A.DefaultRunProperties() {FontSize = 2000};

            level6ParagraphProperties4.Append(defaultRunProperties40);

            var level7ParagraphProperties4 = new A.Level7ParagraphProperties();
            var defaultRunProperties41 = new A.DefaultRunProperties() {FontSize = 2000};

            level7ParagraphProperties4.Append(defaultRunProperties41);

            var level8ParagraphProperties4 = new A.Level8ParagraphProperties();
            var defaultRunProperties42 = new A.DefaultRunProperties() {FontSize = 2000};

            level8ParagraphProperties4.Append(defaultRunProperties42);

            var level9ParagraphProperties4 = new A.Level9ParagraphProperties();
            var defaultRunProperties43 = new A.DefaultRunProperties() {FontSize = 2000};

            level9ParagraphProperties4.Append(defaultRunProperties43);

            listStyle12.Append(level1ParagraphProperties9);
            listStyle12.Append(level2ParagraphProperties4);
            listStyle12.Append(level3ParagraphProperties4);
            listStyle12.Append(level4ParagraphProperties4);
            listStyle12.Append(level5ParagraphProperties4);
            listStyle12.Append(level6ParagraphProperties4);
            listStyle12.Append(level7ParagraphProperties4);
            listStyle12.Append(level8ParagraphProperties4);
            listStyle12.Append(level9ParagraphProperties4);

            var paragraph16 = new A.Paragraph();
            var paragraphProperties6 = new A.ParagraphProperties() {Level = 0};

            var run9 = new A.Run();

            var runProperties13 = new A.RunProperties() {Language = "en-US"};
            runProperties13.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text13 = new A.Text();
            text13.Text = "Click to edit Master text styles";

            run9.Append(runProperties13);
            run9.Append(text13);

            paragraph16.Append(paragraphProperties6);
            paragraph16.Append(run9);

            var paragraph17 = new A.Paragraph();
            var paragraphProperties7 = new A.ParagraphProperties() {Level = 1};

            var run10 = new A.Run();

            var runProperties14 = new A.RunProperties() {Language = "en-US"};
            runProperties14.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text14 = new A.Text();
            text14.Text = "Second level";

            run10.Append(runProperties14);
            run10.Append(text14);

            paragraph17.Append(paragraphProperties7);
            paragraph17.Append(run10);

            var paragraph18 = new A.Paragraph();
            var paragraphProperties8 = new A.ParagraphProperties() {Level = 2};

            var run11 = new A.Run();

            var runProperties15 = new A.RunProperties() {Language = "en-US"};
            runProperties15.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text15 = new A.Text();
            text15.Text = "Third level";

            run11.Append(runProperties15);
            run11.Append(text15);

            paragraph18.Append(paragraphProperties8);
            paragraph18.Append(run11);

            var paragraph19 = new A.Paragraph();
            var paragraphProperties9 = new A.ParagraphProperties() {Level = 3};

            var run12 = new A.Run();

            var runProperties16 = new A.RunProperties() {Language = "en-US"};
            runProperties16.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text16 = new A.Text();
            text16.Text = "Fourth level";

            run12.Append(runProperties16);
            run12.Append(text16);

            paragraph19.Append(paragraphProperties9);
            paragraph19.Append(run12);

            var paragraph20 = new A.Paragraph();
            var paragraphProperties10 = new A.ParagraphProperties() {Level = 4};

            var run13 = new A.Run();

            var runProperties17 = new A.RunProperties() {Language = "en-US"};
            runProperties17.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text17 = new A.Text();
            text17.Text = "Fifth level";

            run13.Append(runProperties17);
            run13.Append(text17);
            var endParagraphRunProperties12 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph20.Append(paragraphProperties10);
            paragraph20.Append(run13);
            paragraph20.Append(endParagraphRunProperties12);

            textBody12.Append(bodyProperties12);
            textBody12.Append(listStyle12);
            textBody12.Append(paragraph16);
            textBody12.Append(paragraph17);
            textBody12.Append(paragraph18);
            textBody12.Append(paragraph19);
            textBody12.Append(paragraph20);

            shape12.Append(nonVisualShapeProperties12);
            shape12.Append(shapeProperties12);
            shape12.Append(textBody12);

            var shape13 = new Shape();

            var nonVisualShapeProperties13 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties17 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Text Placeholder 3"
            };

            var nonVisualShapeDrawingProperties13 = new NonVisualShapeDrawingProperties();
            var shapeLocks13 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties13.Append(shapeLocks13);

            var applicationNonVisualDrawingProperties17 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape13 = new PlaceholderShape() {
                Type = PlaceholderValues.Body,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 2U
            };

            applicationNonVisualDrawingProperties17.Append(placeholderShape13);

            nonVisualShapeProperties13.Append(nonVisualDrawingProperties17);
            nonVisualShapeProperties13.Append(nonVisualShapeDrawingProperties13);
            nonVisualShapeProperties13.Append(applicationNonVisualDrawingProperties17);

            var shapeProperties13 = new ShapeProperties();

            var transform2D8 = new A.Transform2D();
            var offset12 = new A.Offset() {X = 457200L, Y = 1435100L};
            var extents12 = new A.Extents() {Cx = 3008313L, Cy = 4691063L};

            transform2D8.Append(offset12);
            transform2D8.Append(extents12);

            shapeProperties13.Append(transform2D8);

            var textBody13 = new TextBody();
            var bodyProperties13 = new A.BodyProperties();

            var listStyle13 = new A.ListStyle();

            var level1ParagraphProperties10 = new A.Level1ParagraphProperties() {LeftMargin = 0, Indent = 0};
            var noBullet2 = new A.NoBullet();
            var defaultRunProperties44 = new A.DefaultRunProperties() {FontSize = 1400};

            level1ParagraphProperties10.Append(noBullet2);
            level1ParagraphProperties10.Append(defaultRunProperties44);

            var level2ParagraphProperties5 = new A.Level2ParagraphProperties() {LeftMargin = 457200, Indent = 0};
            var noBullet3 = new A.NoBullet();
            var defaultRunProperties45 = new A.DefaultRunProperties() {FontSize = 1200};

            level2ParagraphProperties5.Append(noBullet3);
            level2ParagraphProperties5.Append(defaultRunProperties45);

            var level3ParagraphProperties5 = new A.Level3ParagraphProperties() {LeftMargin = 914400, Indent = 0};
            var noBullet4 = new A.NoBullet();
            var defaultRunProperties46 = new A.DefaultRunProperties() {FontSize = 1000};

            level3ParagraphProperties5.Append(noBullet4);
            level3ParagraphProperties5.Append(defaultRunProperties46);

            var level4ParagraphProperties5 = new A.Level4ParagraphProperties() {LeftMargin = 1371600, Indent = 0};
            var noBullet5 = new A.NoBullet();
            var defaultRunProperties47 = new A.DefaultRunProperties() {FontSize = 900};

            level4ParagraphProperties5.Append(noBullet5);
            level4ParagraphProperties5.Append(defaultRunProperties47);

            var level5ParagraphProperties5 = new A.Level5ParagraphProperties() {LeftMargin = 1828800, Indent = 0};
            var noBullet6 = new A.NoBullet();
            var defaultRunProperties48 = new A.DefaultRunProperties() {FontSize = 900};

            level5ParagraphProperties5.Append(noBullet6);
            level5ParagraphProperties5.Append(defaultRunProperties48);

            var level6ParagraphProperties5 = new A.Level6ParagraphProperties() {LeftMargin = 2286000, Indent = 0};
            var noBullet7 = new A.NoBullet();
            var defaultRunProperties49 = new A.DefaultRunProperties() {FontSize = 900};

            level6ParagraphProperties5.Append(noBullet7);
            level6ParagraphProperties5.Append(defaultRunProperties49);

            var level7ParagraphProperties5 = new A.Level7ParagraphProperties() {LeftMargin = 2743200, Indent = 0};
            var noBullet8 = new A.NoBullet();
            var defaultRunProperties50 = new A.DefaultRunProperties() {FontSize = 900};

            level7ParagraphProperties5.Append(noBullet8);
            level7ParagraphProperties5.Append(defaultRunProperties50);

            var level8ParagraphProperties5 = new A.Level8ParagraphProperties() {LeftMargin = 3200400, Indent = 0};
            var noBullet9 = new A.NoBullet();
            var defaultRunProperties51 = new A.DefaultRunProperties() {FontSize = 900};

            level8ParagraphProperties5.Append(noBullet9);
            level8ParagraphProperties5.Append(defaultRunProperties51);

            var level9ParagraphProperties5 = new A.Level9ParagraphProperties() {LeftMargin = 3657600, Indent = 0};
            var noBullet10 = new A.NoBullet();
            var defaultRunProperties52 = new A.DefaultRunProperties() {FontSize = 900};

            level9ParagraphProperties5.Append(noBullet10);
            level9ParagraphProperties5.Append(defaultRunProperties52);

            listStyle13.Append(level1ParagraphProperties10);
            listStyle13.Append(level2ParagraphProperties5);
            listStyle13.Append(level3ParagraphProperties5);
            listStyle13.Append(level4ParagraphProperties5);
            listStyle13.Append(level5ParagraphProperties5);
            listStyle13.Append(level6ParagraphProperties5);
            listStyle13.Append(level7ParagraphProperties5);
            listStyle13.Append(level8ParagraphProperties5);
            listStyle13.Append(level9ParagraphProperties5);

            var paragraph21 = new A.Paragraph();
            var paragraphProperties11 = new A.ParagraphProperties() {Level = 0};

            var run14 = new A.Run();

            var runProperties18 = new A.RunProperties() {Language = "en-US"};
            runProperties18.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text18 = new A.Text();
            text18.Text = "Click to edit Master text styles";

            run14.Append(runProperties18);
            run14.Append(text18);

            paragraph21.Append(paragraphProperties11);
            paragraph21.Append(run14);

            textBody13.Append(bodyProperties13);
            textBody13.Append(listStyle13);
            textBody13.Append(paragraph21);

            shape13.Append(nonVisualShapeProperties13);
            shape13.Append(shapeProperties13);
            shape13.Append(textBody13);

            var shape14 = new Shape();

            var nonVisualShapeProperties14 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties18 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 5U,
                Name = "Date Placeholder 4"
            };

            var nonVisualShapeDrawingProperties14 = new NonVisualShapeDrawingProperties();
            var shapeLocks14 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties14.Append(shapeLocks14);

            var applicationNonVisualDrawingProperties18 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape14 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 10U
            };

            applicationNonVisualDrawingProperties18.Append(placeholderShape14);

            nonVisualShapeProperties14.Append(nonVisualDrawingProperties18);
            nonVisualShapeProperties14.Append(nonVisualShapeDrawingProperties14);
            nonVisualShapeProperties14.Append(applicationNonVisualDrawingProperties18);
            var shapeProperties14 = new ShapeProperties();

            var textBody14 = new TextBody();
            var bodyProperties14 = new A.BodyProperties();
            var listStyle14 = new A.ListStyle();

            var paragraph22 = new A.Paragraph();

            var field5 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties19 = new A.RunProperties() {Language = "en-GB"};
            runProperties19.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text19 = new A.Text();
            text19.Text = "04/06/2014";

            field5.Append(runProperties19);
            field5.Append(text19);
            var endParagraphRunProperties13 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph22.Append(field5);
            paragraph22.Append(endParagraphRunProperties13);

            textBody14.Append(bodyProperties14);
            textBody14.Append(listStyle14);
            textBody14.Append(paragraph22);

            shape14.Append(nonVisualShapeProperties14);
            shape14.Append(shapeProperties14);
            shape14.Append(textBody14);

            var shape15 = new Shape();

            var nonVisualShapeProperties15 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties19 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 6U,
                Name = "Footer Placeholder 5"
            };

            var nonVisualShapeDrawingProperties15 = new NonVisualShapeDrawingProperties();
            var shapeLocks15 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties15.Append(shapeLocks15);

            var applicationNonVisualDrawingProperties19 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape15 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 11U
            };

            applicationNonVisualDrawingProperties19.Append(placeholderShape15);

            nonVisualShapeProperties15.Append(nonVisualDrawingProperties19);
            nonVisualShapeProperties15.Append(nonVisualShapeDrawingProperties15);
            nonVisualShapeProperties15.Append(applicationNonVisualDrawingProperties19);
            var shapeProperties15 = new ShapeProperties();

            var textBody15 = new TextBody();
            var bodyProperties15 = new A.BodyProperties();
            var listStyle15 = new A.ListStyle();

            var paragraph23 = new A.Paragraph();
            var endParagraphRunProperties14 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph23.Append(endParagraphRunProperties14);

            textBody15.Append(bodyProperties15);
            textBody15.Append(listStyle15);
            textBody15.Append(paragraph23);

            shape15.Append(nonVisualShapeProperties15);
            shape15.Append(shapeProperties15);
            shape15.Append(textBody15);

            var shape16 = new Shape();

            var nonVisualShapeProperties16 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties20 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 7U,
                Name = "Slide Number Placeholder 6"
            };

            var nonVisualShapeDrawingProperties16 = new NonVisualShapeDrawingProperties();
            var shapeLocks16 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties16.Append(shapeLocks16);

            var applicationNonVisualDrawingProperties20 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape16 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 12U
            };

            applicationNonVisualDrawingProperties20.Append(placeholderShape16);

            nonVisualShapeProperties16.Append(nonVisualDrawingProperties20);
            nonVisualShapeProperties16.Append(nonVisualShapeDrawingProperties16);
            nonVisualShapeProperties16.Append(applicationNonVisualDrawingProperties20);
            var shapeProperties16 = new ShapeProperties();

            var textBody16 = new TextBody();
            var bodyProperties16 = new A.BodyProperties();
            var listStyle16 = new A.ListStyle();

            var paragraph24 = new A.Paragraph();

            var field6 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties20 = new A.RunProperties() {Language = "en-GB"};
            runProperties20.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text20 = new A.Text();
            text20.Text = "‹#›";

            field6.Append(runProperties20);
            field6.Append(text20);
            var endParagraphRunProperties15 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph24.Append(field6);
            paragraph24.Append(endParagraphRunProperties15);

            textBody16.Append(bodyProperties16);
            textBody16.Append(listStyle16);
            textBody16.Append(paragraph24);

            shape16.Append(nonVisualShapeProperties16);
            shape16.Append(shapeProperties16);
            shape16.Append(textBody16);

            shapeTree4.Append(nonVisualGroupShapeProperties4);
            shapeTree4.Append(groupShapeProperties4);
            shapeTree4.Append(shape11);
            shapeTree4.Append(shape12);
            shapeTree4.Append(shape13);
            shapeTree4.Append(shape14);
            shapeTree4.Append(shape15);
            shapeTree4.Append(shape16);

            var commonSlideDataExtensionList4 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension4 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId4 = new P14.CreationId() {Val = (UInt32Value) 3469623364U};
            creationId4.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension4.Append(creationId4);

            commonSlideDataExtensionList4.Append(commonSlideDataExtension4);

            commonSlideData4.Append(shapeTree4);
            commonSlideData4.Append(commonSlideDataExtensionList4);

            var colorMapOverride3 = new ColorMapOverride();
            var masterColorMapping3 = new A.MasterColorMapping();

            colorMapOverride3.Append(masterColorMapping3);

            slideLayout2.Append(commonSlideData4);
            slideLayout2.Append(colorMapOverride3);

            slideLayoutPart2.SlideLayout = slideLayout2;
        }

        // Generates content of slideLayoutPart3.
        private static void GenerateSlideLayoutPart3Content(SlideLayoutPart slideLayoutPart3)
        {
            var slideLayout3 = new SlideLayout() {Type = SlideLayoutValues.SectionHeader, Preserve = true};
            slideLayout3.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout3.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout3.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData5 = new CommonSlideData() {Name = "Section Header"};

            var shapeTree5 = new ShapeTree();

            var nonVisualGroupShapeProperties5 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties21 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties5 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties21 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties5.Append(nonVisualDrawingProperties21);
            nonVisualGroupShapeProperties5.Append(nonVisualGroupShapeDrawingProperties5);
            nonVisualGroupShapeProperties5.Append(applicationNonVisualDrawingProperties21);

            var groupShapeProperties5 = new GroupShapeProperties();

            var transformGroup5 = new A.TransformGroup();
            var offset13 = new A.Offset() {X = 0L, Y = 0L};
            var extents13 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset5 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents5 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup5.Append(offset13);
            transformGroup5.Append(extents13);
            transformGroup5.Append(childOffset5);
            transformGroup5.Append(childExtents5);

            groupShapeProperties5.Append(transformGroup5);

            var shape17 = new Shape();

            var nonVisualShapeProperties17 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties22 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 2U,
                Name = "Title 1"
            };

            var nonVisualShapeDrawingProperties17 = new NonVisualShapeDrawingProperties();
            var shapeLocks17 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties17.Append(shapeLocks17);

            var applicationNonVisualDrawingProperties22 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape17 = new PlaceholderShape() {Type = PlaceholderValues.Title};

            applicationNonVisualDrawingProperties22.Append(placeholderShape17);

            nonVisualShapeProperties17.Append(nonVisualDrawingProperties22);
            nonVisualShapeProperties17.Append(nonVisualShapeDrawingProperties17);
            nonVisualShapeProperties17.Append(applicationNonVisualDrawingProperties22);

            var shapeProperties17 = new ShapeProperties();

            var transform2D9 = new A.Transform2D();
            var offset14 = new A.Offset() {X = 722313L, Y = 4406900L};
            var extents14 = new A.Extents() {Cx = 7772400L, Cy = 1362075L};

            transform2D9.Append(offset14);
            transform2D9.Append(extents14);

            shapeProperties17.Append(transform2D9);

            var textBody17 = new TextBody();
            var bodyProperties17 = new A.BodyProperties() {Anchor = A.TextAnchoringTypeValues.Top};

            var listStyle17 = new A.ListStyle();

            var level1ParagraphProperties11 = new A.Level1ParagraphProperties() {
                Alignment = A.TextAlignmentTypeValues.Left
            };
            var defaultRunProperties53 = new A.DefaultRunProperties() {
                FontSize = 4000,
                Bold = true,
                Capital = A.TextCapsValues.All
            };

            level1ParagraphProperties11.Append(defaultRunProperties53);

            listStyle17.Append(level1ParagraphProperties11);

            var paragraph25 = new A.Paragraph();

            var run15 = new A.Run();

            var runProperties21 = new A.RunProperties() {Language = "en-US"};
            runProperties21.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text21 = new A.Text();
            text21.Text = "Click to edit Master title style";

            run15.Append(runProperties21);
            run15.Append(text21);
            var endParagraphRunProperties16 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph25.Append(run15);
            paragraph25.Append(endParagraphRunProperties16);

            textBody17.Append(bodyProperties17);
            textBody17.Append(listStyle17);
            textBody17.Append(paragraph25);

            shape17.Append(nonVisualShapeProperties17);
            shape17.Append(shapeProperties17);
            shape17.Append(textBody17);

            var shape18 = new Shape();

            var nonVisualShapeProperties18 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties23 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Text Placeholder 2"
            };

            var nonVisualShapeDrawingProperties18 = new NonVisualShapeDrawingProperties();
            var shapeLocks18 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties18.Append(shapeLocks18);

            var applicationNonVisualDrawingProperties23 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape18 = new PlaceholderShape() {Type = PlaceholderValues.Body, Index = (UInt32Value) 1U};

            applicationNonVisualDrawingProperties23.Append(placeholderShape18);

            nonVisualShapeProperties18.Append(nonVisualDrawingProperties23);
            nonVisualShapeProperties18.Append(nonVisualShapeDrawingProperties18);
            nonVisualShapeProperties18.Append(applicationNonVisualDrawingProperties23);

            var shapeProperties18 = new ShapeProperties();

            var transform2D10 = new A.Transform2D();
            var offset15 = new A.Offset() {X = 722313L, Y = 2906713L};
            var extents15 = new A.Extents() {Cx = 7772400L, Cy = 1500187L};

            transform2D10.Append(offset15);
            transform2D10.Append(extents15);

            shapeProperties18.Append(transform2D10);

            var textBody18 = new TextBody();
            var bodyProperties18 = new A.BodyProperties() {Anchor = A.TextAnchoringTypeValues.Bottom};

            var listStyle18 = new A.ListStyle();

            var level1ParagraphProperties12 = new A.Level1ParagraphProperties() {LeftMargin = 0, Indent = 0};
            var noBullet11 = new A.NoBullet();

            var defaultRunProperties54 = new A.DefaultRunProperties() {FontSize = 2000};

            var solidFill32 = new A.SolidFill();

            var schemeColor33 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint4 = new A.Tint() {Val = 75000};

            schemeColor33.Append(tint4);

            solidFill32.Append(schemeColor33);

            defaultRunProperties54.Append(solidFill32);

            level1ParagraphProperties12.Append(noBullet11);
            level1ParagraphProperties12.Append(defaultRunProperties54);

            var level2ParagraphProperties6 = new A.Level2ParagraphProperties() {LeftMargin = 457200, Indent = 0};
            var noBullet12 = new A.NoBullet();

            var defaultRunProperties55 = new A.DefaultRunProperties() {FontSize = 1800};

            var solidFill33 = new A.SolidFill();

            var schemeColor34 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint5 = new A.Tint() {Val = 75000};

            schemeColor34.Append(tint5);

            solidFill33.Append(schemeColor34);

            defaultRunProperties55.Append(solidFill33);

            level2ParagraphProperties6.Append(noBullet12);
            level2ParagraphProperties6.Append(defaultRunProperties55);

            var level3ParagraphProperties6 = new A.Level3ParagraphProperties() {LeftMargin = 914400, Indent = 0};
            var noBullet13 = new A.NoBullet();

            var defaultRunProperties56 = new A.DefaultRunProperties() {FontSize = 1600};

            var solidFill34 = new A.SolidFill();

            var schemeColor35 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint6 = new A.Tint() {Val = 75000};

            schemeColor35.Append(tint6);

            solidFill34.Append(schemeColor35);

            defaultRunProperties56.Append(solidFill34);

            level3ParagraphProperties6.Append(noBullet13);
            level3ParagraphProperties6.Append(defaultRunProperties56);

            var level4ParagraphProperties6 = new A.Level4ParagraphProperties() {LeftMargin = 1371600, Indent = 0};
            var noBullet14 = new A.NoBullet();

            var defaultRunProperties57 = new A.DefaultRunProperties() {FontSize = 1400};

            var solidFill35 = new A.SolidFill();

            var schemeColor36 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint7 = new A.Tint() {Val = 75000};

            schemeColor36.Append(tint7);

            solidFill35.Append(schemeColor36);

            defaultRunProperties57.Append(solidFill35);

            level4ParagraphProperties6.Append(noBullet14);
            level4ParagraphProperties6.Append(defaultRunProperties57);

            var level5ParagraphProperties6 = new A.Level5ParagraphProperties() {LeftMargin = 1828800, Indent = 0};
            var noBullet15 = new A.NoBullet();

            var defaultRunProperties58 = new A.DefaultRunProperties() {FontSize = 1400};

            var solidFill36 = new A.SolidFill();

            var schemeColor37 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint8 = new A.Tint() {Val = 75000};

            schemeColor37.Append(tint8);

            solidFill36.Append(schemeColor37);

            defaultRunProperties58.Append(solidFill36);

            level5ParagraphProperties6.Append(noBullet15);
            level5ParagraphProperties6.Append(defaultRunProperties58);

            var level6ParagraphProperties6 = new A.Level6ParagraphProperties() {LeftMargin = 2286000, Indent = 0};
            var noBullet16 = new A.NoBullet();

            var defaultRunProperties59 = new A.DefaultRunProperties() {FontSize = 1400};

            var solidFill37 = new A.SolidFill();

            var schemeColor38 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint9 = new A.Tint() {Val = 75000};

            schemeColor38.Append(tint9);

            solidFill37.Append(schemeColor38);

            defaultRunProperties59.Append(solidFill37);

            level6ParagraphProperties6.Append(noBullet16);
            level6ParagraphProperties6.Append(defaultRunProperties59);

            var level7ParagraphProperties6 = new A.Level7ParagraphProperties() {LeftMargin = 2743200, Indent = 0};
            var noBullet17 = new A.NoBullet();

            var defaultRunProperties60 = new A.DefaultRunProperties() {FontSize = 1400};

            var solidFill38 = new A.SolidFill();

            var schemeColor39 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint10 = new A.Tint() {Val = 75000};

            schemeColor39.Append(tint10);

            solidFill38.Append(schemeColor39);

            defaultRunProperties60.Append(solidFill38);

            level7ParagraphProperties6.Append(noBullet17);
            level7ParagraphProperties6.Append(defaultRunProperties60);

            var level8ParagraphProperties6 = new A.Level8ParagraphProperties() {LeftMargin = 3200400, Indent = 0};
            var noBullet18 = new A.NoBullet();

            var defaultRunProperties61 = new A.DefaultRunProperties() {FontSize = 1400};

            var solidFill39 = new A.SolidFill();

            var schemeColor40 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint11 = new A.Tint() {Val = 75000};

            schemeColor40.Append(tint11);

            solidFill39.Append(schemeColor40);

            defaultRunProperties61.Append(solidFill39);

            level8ParagraphProperties6.Append(noBullet18);
            level8ParagraphProperties6.Append(defaultRunProperties61);

            var level9ParagraphProperties6 = new A.Level9ParagraphProperties() {LeftMargin = 3657600, Indent = 0};
            var noBullet19 = new A.NoBullet();

            var defaultRunProperties62 = new A.DefaultRunProperties() {FontSize = 1400};

            var solidFill40 = new A.SolidFill();

            var schemeColor41 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint12 = new A.Tint() {Val = 75000};

            schemeColor41.Append(tint12);

            solidFill40.Append(schemeColor41);

            defaultRunProperties62.Append(solidFill40);

            level9ParagraphProperties6.Append(noBullet19);
            level9ParagraphProperties6.Append(defaultRunProperties62);

            listStyle18.Append(level1ParagraphProperties12);
            listStyle18.Append(level2ParagraphProperties6);
            listStyle18.Append(level3ParagraphProperties6);
            listStyle18.Append(level4ParagraphProperties6);
            listStyle18.Append(level5ParagraphProperties6);
            listStyle18.Append(level6ParagraphProperties6);
            listStyle18.Append(level7ParagraphProperties6);
            listStyle18.Append(level8ParagraphProperties6);
            listStyle18.Append(level9ParagraphProperties6);

            var paragraph26 = new A.Paragraph();
            var paragraphProperties12 = new A.ParagraphProperties() {Level = 0};

            var run16 = new A.Run();

            var runProperties22 = new A.RunProperties() {Language = "en-US"};
            runProperties22.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text22 = new A.Text();
            text22.Text = "Click to edit Master text styles";

            run16.Append(runProperties22);
            run16.Append(text22);

            paragraph26.Append(paragraphProperties12);
            paragraph26.Append(run16);

            textBody18.Append(bodyProperties18);
            textBody18.Append(listStyle18);
            textBody18.Append(paragraph26);

            shape18.Append(nonVisualShapeProperties18);
            shape18.Append(shapeProperties18);
            shape18.Append(textBody18);

            var shape19 = new Shape();

            var nonVisualShapeProperties19 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties24 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Date Placeholder 3"
            };

            var nonVisualShapeDrawingProperties19 = new NonVisualShapeDrawingProperties();
            var shapeLocks19 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties19.Append(shapeLocks19);

            var applicationNonVisualDrawingProperties24 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape19 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 10U
            };

            applicationNonVisualDrawingProperties24.Append(placeholderShape19);

            nonVisualShapeProperties19.Append(nonVisualDrawingProperties24);
            nonVisualShapeProperties19.Append(nonVisualShapeDrawingProperties19);
            nonVisualShapeProperties19.Append(applicationNonVisualDrawingProperties24);
            var shapeProperties19 = new ShapeProperties();

            var textBody19 = new TextBody();
            var bodyProperties19 = new A.BodyProperties();
            var listStyle19 = new A.ListStyle();

            var paragraph27 = new A.Paragraph();

            var field7 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties23 = new A.RunProperties() {Language = "en-GB"};
            runProperties23.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text23 = new A.Text();
            text23.Text = "04/06/2014";

            field7.Append(runProperties23);
            field7.Append(text23);
            var endParagraphRunProperties17 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph27.Append(field7);
            paragraph27.Append(endParagraphRunProperties17);

            textBody19.Append(bodyProperties19);
            textBody19.Append(listStyle19);
            textBody19.Append(paragraph27);

            shape19.Append(nonVisualShapeProperties19);
            shape19.Append(shapeProperties19);
            shape19.Append(textBody19);

            var shape20 = new Shape();

            var nonVisualShapeProperties20 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties25 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 5U,
                Name = "Footer Placeholder 4"
            };

            var nonVisualShapeDrawingProperties20 = new NonVisualShapeDrawingProperties();
            var shapeLocks20 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties20.Append(shapeLocks20);

            var applicationNonVisualDrawingProperties25 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape20 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 11U
            };

            applicationNonVisualDrawingProperties25.Append(placeholderShape20);

            nonVisualShapeProperties20.Append(nonVisualDrawingProperties25);
            nonVisualShapeProperties20.Append(nonVisualShapeDrawingProperties20);
            nonVisualShapeProperties20.Append(applicationNonVisualDrawingProperties25);
            var shapeProperties20 = new ShapeProperties();

            var textBody20 = new TextBody();
            var bodyProperties20 = new A.BodyProperties();
            var listStyle20 = new A.ListStyle();

            var paragraph28 = new A.Paragraph();
            var endParagraphRunProperties18 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph28.Append(endParagraphRunProperties18);

            textBody20.Append(bodyProperties20);
            textBody20.Append(listStyle20);
            textBody20.Append(paragraph28);

            shape20.Append(nonVisualShapeProperties20);
            shape20.Append(shapeProperties20);
            shape20.Append(textBody20);

            var shape21 = new Shape();

            var nonVisualShapeProperties21 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties26 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 6U,
                Name = "Slide Number Placeholder 5"
            };

            var nonVisualShapeDrawingProperties21 = new NonVisualShapeDrawingProperties();
            var shapeLocks21 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties21.Append(shapeLocks21);

            var applicationNonVisualDrawingProperties26 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape21 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 12U
            };

            applicationNonVisualDrawingProperties26.Append(placeholderShape21);

            nonVisualShapeProperties21.Append(nonVisualDrawingProperties26);
            nonVisualShapeProperties21.Append(nonVisualShapeDrawingProperties21);
            nonVisualShapeProperties21.Append(applicationNonVisualDrawingProperties26);
            var shapeProperties21 = new ShapeProperties();

            var textBody21 = new TextBody();
            var bodyProperties21 = new A.BodyProperties();
            var listStyle21 = new A.ListStyle();

            var paragraph29 = new A.Paragraph();

            var field8 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties24 = new A.RunProperties() {Language = "en-GB"};
            runProperties24.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text24 = new A.Text();
            text24.Text = "‹#›";

            field8.Append(runProperties24);
            field8.Append(text24);
            var endParagraphRunProperties19 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph29.Append(field8);
            paragraph29.Append(endParagraphRunProperties19);

            textBody21.Append(bodyProperties21);
            textBody21.Append(listStyle21);
            textBody21.Append(paragraph29);

            shape21.Append(nonVisualShapeProperties21);
            shape21.Append(shapeProperties21);
            shape21.Append(textBody21);

            shapeTree5.Append(nonVisualGroupShapeProperties5);
            shapeTree5.Append(groupShapeProperties5);
            shapeTree5.Append(shape17);
            shapeTree5.Append(shape18);
            shapeTree5.Append(shape19);
            shapeTree5.Append(shape20);
            shapeTree5.Append(shape21);

            var commonSlideDataExtensionList5 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension5 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId5 = new P14.CreationId() {Val = (UInt32Value) 372079535U};
            creationId5.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension5.Append(creationId5);

            commonSlideDataExtensionList5.Append(commonSlideDataExtension5);

            commonSlideData5.Append(shapeTree5);
            commonSlideData5.Append(commonSlideDataExtensionList5);

            var colorMapOverride4 = new ColorMapOverride();
            var masterColorMapping4 = new A.MasterColorMapping();

            colorMapOverride4.Append(masterColorMapping4);

            slideLayout3.Append(commonSlideData5);
            slideLayout3.Append(colorMapOverride4);

            slideLayoutPart3.SlideLayout = slideLayout3;
        }

        // Generates content of slideLayoutPart4.
        private static void GenerateSlideLayoutPart4Content(SlideLayoutPart slideLayoutPart4)
        {
            var slideLayout4 = new SlideLayout() {Type = SlideLayoutValues.Blank, Preserve = true};
            slideLayout4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout4.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout4.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData6 = new CommonSlideData() {Name = "Blank"};

            var shapeTree6 = new ShapeTree();

            var nonVisualGroupShapeProperties6 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties27 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties6 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties27 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties6.Append(nonVisualDrawingProperties27);
            nonVisualGroupShapeProperties6.Append(nonVisualGroupShapeDrawingProperties6);
            nonVisualGroupShapeProperties6.Append(applicationNonVisualDrawingProperties27);

            var groupShapeProperties6 = new GroupShapeProperties();

            var transformGroup6 = new A.TransformGroup();
            var offset16 = new A.Offset() {X = 0L, Y = 0L};
            var extents16 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset6 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents6 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup6.Append(offset16);
            transformGroup6.Append(extents16);
            transformGroup6.Append(childOffset6);
            transformGroup6.Append(childExtents6);

            groupShapeProperties6.Append(transformGroup6);

            var shape22 = new Shape();

            var nonVisualShapeProperties22 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties28 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 2U,
                Name = "Date Placeholder 1"
            };

            var nonVisualShapeDrawingProperties22 = new NonVisualShapeDrawingProperties();
            var shapeLocks22 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties22.Append(shapeLocks22);

            var applicationNonVisualDrawingProperties28 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape22 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 10U
            };

            applicationNonVisualDrawingProperties28.Append(placeholderShape22);

            nonVisualShapeProperties22.Append(nonVisualDrawingProperties28);
            nonVisualShapeProperties22.Append(nonVisualShapeDrawingProperties22);
            nonVisualShapeProperties22.Append(applicationNonVisualDrawingProperties28);
            var shapeProperties22 = new ShapeProperties();

            var textBody22 = new TextBody();
            var bodyProperties22 = new A.BodyProperties();
            var listStyle22 = new A.ListStyle();

            var paragraph30 = new A.Paragraph();

            var field9 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties25 = new A.RunProperties() {Language = "en-GB"};
            runProperties25.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text25 = new A.Text();
            text25.Text = "04/06/2014";

            field9.Append(runProperties25);
            field9.Append(text25);
            var endParagraphRunProperties20 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph30.Append(field9);
            paragraph30.Append(endParagraphRunProperties20);

            textBody22.Append(bodyProperties22);
            textBody22.Append(listStyle22);
            textBody22.Append(paragraph30);

            shape22.Append(nonVisualShapeProperties22);
            shape22.Append(shapeProperties22);
            shape22.Append(textBody22);

            var shape23 = new Shape();

            var nonVisualShapeProperties23 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties29 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Footer Placeholder 2"
            };

            var nonVisualShapeDrawingProperties23 = new NonVisualShapeDrawingProperties();
            var shapeLocks23 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties23.Append(shapeLocks23);

            var applicationNonVisualDrawingProperties29 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape23 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 11U
            };

            applicationNonVisualDrawingProperties29.Append(placeholderShape23);

            nonVisualShapeProperties23.Append(nonVisualDrawingProperties29);
            nonVisualShapeProperties23.Append(nonVisualShapeDrawingProperties23);
            nonVisualShapeProperties23.Append(applicationNonVisualDrawingProperties29);
            var shapeProperties23 = new ShapeProperties();

            var textBody23 = new TextBody();
            var bodyProperties23 = new A.BodyProperties();
            var listStyle23 = new A.ListStyle();

            var paragraph31 = new A.Paragraph();
            var endParagraphRunProperties21 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph31.Append(endParagraphRunProperties21);

            textBody23.Append(bodyProperties23);
            textBody23.Append(listStyle23);
            textBody23.Append(paragraph31);

            shape23.Append(nonVisualShapeProperties23);
            shape23.Append(shapeProperties23);
            shape23.Append(textBody23);

            var shape24 = new Shape();

            var nonVisualShapeProperties24 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties30 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Slide Number Placeholder 3"
            };

            var nonVisualShapeDrawingProperties24 = new NonVisualShapeDrawingProperties();
            var shapeLocks24 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties24.Append(shapeLocks24);

            var applicationNonVisualDrawingProperties30 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape24 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 12U
            };

            applicationNonVisualDrawingProperties30.Append(placeholderShape24);

            nonVisualShapeProperties24.Append(nonVisualDrawingProperties30);
            nonVisualShapeProperties24.Append(nonVisualShapeDrawingProperties24);
            nonVisualShapeProperties24.Append(applicationNonVisualDrawingProperties30);
            var shapeProperties24 = new ShapeProperties();

            var textBody24 = new TextBody();
            var bodyProperties24 = new A.BodyProperties();
            var listStyle24 = new A.ListStyle();

            var paragraph32 = new A.Paragraph();

            var field10 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties26 = new A.RunProperties() {Language = "en-GB"};
            runProperties26.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text26 = new A.Text();
            text26.Text = "‹#›";

            field10.Append(runProperties26);
            field10.Append(text26);
            var endParagraphRunProperties22 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph32.Append(field10);
            paragraph32.Append(endParagraphRunProperties22);

            textBody24.Append(bodyProperties24);
            textBody24.Append(listStyle24);
            textBody24.Append(paragraph32);

            shape24.Append(nonVisualShapeProperties24);
            shape24.Append(shapeProperties24);
            shape24.Append(textBody24);

            shapeTree6.Append(nonVisualGroupShapeProperties6);
            shapeTree6.Append(groupShapeProperties6);
            shapeTree6.Append(shape22);
            shapeTree6.Append(shape23);
            shapeTree6.Append(shape24);

            var commonSlideDataExtensionList6 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension6 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId6 = new P14.CreationId() {Val = (UInt32Value) 1812897738U};
            creationId6.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension6.Append(creationId6);

            commonSlideDataExtensionList6.Append(commonSlideDataExtension6);

            commonSlideData6.Append(shapeTree6);
            commonSlideData6.Append(commonSlideDataExtensionList6);

            var colorMapOverride5 = new ColorMapOverride();
            var masterColorMapping5 = new A.MasterColorMapping();

            colorMapOverride5.Append(masterColorMapping5);

            slideLayout4.Append(commonSlideData6);
            slideLayout4.Append(colorMapOverride5);

            slideLayoutPart4.SlideLayout = slideLayout4;
        }

        // Generates content of themePart1.
        private static void GenerateThemePart1Content(ThemePart themePart1)
        {
            var theme1 = new A.Theme() {Name = "Office Theme"};
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var themeElements1 = new A.ThemeElements();

            var colorScheme1 = new A.ColorScheme() {Name = "Office"};

            var dark1Color1 = new A.Dark1Color();
            var systemColor1 = new A.SystemColor() {Val = A.SystemColorValues.WindowText, LastColor = "000000"};

            dark1Color1.Append(systemColor1);

            var light1Color1 = new A.Light1Color();
            var systemColor2 = new A.SystemColor() {Val = A.SystemColorValues.Window, LastColor = "FFFFFF"};

            light1Color1.Append(systemColor2);

            var dark2Color1 = new A.Dark2Color();
            var rgbColorModelHex1 = new A.RgbColorModelHex() {Val = "1F497D"};

            dark2Color1.Append(rgbColorModelHex1);

            var light2Color1 = new A.Light2Color();
            var rgbColorModelHex2 = new A.RgbColorModelHex() {Val = "EEECE1"};

            light2Color1.Append(rgbColorModelHex2);

            var accent1Color1 = new A.Accent1Color();
            var rgbColorModelHex3 = new A.RgbColorModelHex() {Val = "4F81BD"};

            accent1Color1.Append(rgbColorModelHex3);

            var accent2Color1 = new A.Accent2Color();
            var rgbColorModelHex4 = new A.RgbColorModelHex() {Val = "C0504D"};

            accent2Color1.Append(rgbColorModelHex4);

            var accent3Color1 = new A.Accent3Color();
            var rgbColorModelHex5 = new A.RgbColorModelHex() {Val = "9BBB59"};

            accent3Color1.Append(rgbColorModelHex5);

            var accent4Color1 = new A.Accent4Color();
            var rgbColorModelHex6 = new A.RgbColorModelHex() {Val = "8064A2"};

            accent4Color1.Append(rgbColorModelHex6);

            var accent5Color1 = new A.Accent5Color();
            var rgbColorModelHex7 = new A.RgbColorModelHex() {Val = "4BACC6"};

            accent5Color1.Append(rgbColorModelHex7);

            var accent6Color1 = new A.Accent6Color();
            var rgbColorModelHex8 = new A.RgbColorModelHex() {Val = "F79646"};

            accent6Color1.Append(rgbColorModelHex8);

            var hyperlink1 = new A.Hyperlink();
            var rgbColorModelHex9 = new A.RgbColorModelHex() {Val = "0000FF"};

            hyperlink1.Append(rgbColorModelHex9);

            var followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            var rgbColorModelHex10 = new A.RgbColorModelHex() {Val = "800080"};

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            var fontScheme1 = new A.FontScheme() {Name = "Office"};

            var majorFont1 = new A.MajorFont();
            var latinFont29 = new A.LatinFont() {Typeface = "Calibri"};
            var eastAsianFont29 = new A.EastAsianFont() {Typeface = ""};
            var complexScriptFont29 = new A.ComplexScriptFont() {Typeface = ""};
            var supplementalFont1 = new A.SupplementalFont() {Script = "Jpan", Typeface = "ＭＳ Ｐゴシック"};
            var supplementalFont2 = new A.SupplementalFont() {Script = "Hang", Typeface = "맑은 고딕"};
            var supplementalFont3 = new A.SupplementalFont() {Script = "Hans", Typeface = "宋体"};
            var supplementalFont4 = new A.SupplementalFont() {Script = "Hant", Typeface = "新細明體"};
            var supplementalFont5 = new A.SupplementalFont() {Script = "Arab", Typeface = "Times New Roman"};
            var supplementalFont6 = new A.SupplementalFont() {Script = "Hebr", Typeface = "Times New Roman"};
            var supplementalFont7 = new A.SupplementalFont() {Script = "Thai", Typeface = "Angsana New"};
            var supplementalFont8 = new A.SupplementalFont() {Script = "Ethi", Typeface = "Nyala"};
            var supplementalFont9 = new A.SupplementalFont() {Script = "Beng", Typeface = "Vrinda"};
            var supplementalFont10 = new A.SupplementalFont() {Script = "Gujr", Typeface = "Shruti"};
            var supplementalFont11 = new A.SupplementalFont() {Script = "Khmr", Typeface = "MoolBoran"};
            var supplementalFont12 = new A.SupplementalFont() {Script = "Knda", Typeface = "Tunga"};
            var supplementalFont13 = new A.SupplementalFont() {Script = "Guru", Typeface = "Raavi"};
            var supplementalFont14 = new A.SupplementalFont() {Script = "Cans", Typeface = "Euphemia"};
            var supplementalFont15 = new A.SupplementalFont() {Script = "Cher", Typeface = "Plantagenet Cherokee"};
            var supplementalFont16 = new A.SupplementalFont() {Script = "Yiii", Typeface = "Microsoft Yi Baiti"};
            var supplementalFont17 = new A.SupplementalFont() {Script = "Tibt", Typeface = "Microsoft Himalaya"};
            var supplementalFont18 = new A.SupplementalFont() {Script = "Thaa", Typeface = "MV Boli"};
            var supplementalFont19 = new A.SupplementalFont() {Script = "Deva", Typeface = "Mangal"};
            var supplementalFont20 = new A.SupplementalFont() {Script = "Telu", Typeface = "Gautami"};
            var supplementalFont21 = new A.SupplementalFont() {Script = "Taml", Typeface = "Latha"};
            var supplementalFont22 = new A.SupplementalFont() {Script = "Syrc", Typeface = "Estrangelo Edessa"};
            var supplementalFont23 = new A.SupplementalFont() {Script = "Orya", Typeface = "Kalinga"};
            var supplementalFont24 = new A.SupplementalFont() {Script = "Mlym", Typeface = "Kartika"};
            var supplementalFont25 = new A.SupplementalFont() {Script = "Laoo", Typeface = "DokChampa"};
            var supplementalFont26 = new A.SupplementalFont() {Script = "Sinh", Typeface = "Iskoola Pota"};
            var supplementalFont27 = new A.SupplementalFont() {Script = "Mong", Typeface = "Mongolian Baiti"};
            var supplementalFont28 = new A.SupplementalFont() {Script = "Viet", Typeface = "Times New Roman"};
            var supplementalFont29 = new A.SupplementalFont() {Script = "Uigh", Typeface = "Microsoft Uighur"};
            var supplementalFont30 = new A.SupplementalFont() {Script = "Geor", Typeface = "Sylfaen"};

            majorFont1.Append(latinFont29);
            majorFont1.Append(eastAsianFont29);
            majorFont1.Append(complexScriptFont29);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            var minorFont1 = new A.MinorFont();
            var latinFont30 = new A.LatinFont() {Typeface = "Calibri"};
            var eastAsianFont30 = new A.EastAsianFont() {Typeface = ""};
            var complexScriptFont30 = new A.ComplexScriptFont() {Typeface = ""};
            var supplementalFont31 = new A.SupplementalFont() {Script = "Jpan", Typeface = "ＭＳ Ｐゴシック"};
            var supplementalFont32 = new A.SupplementalFont() {Script = "Hang", Typeface = "맑은 고딕"};
            var supplementalFont33 = new A.SupplementalFont() {Script = "Hans", Typeface = "宋体"};
            var supplementalFont34 = new A.SupplementalFont() {Script = "Hant", Typeface = "新細明體"};
            var supplementalFont35 = new A.SupplementalFont() {Script = "Arab", Typeface = "Arial"};
            var supplementalFont36 = new A.SupplementalFont() {Script = "Hebr", Typeface = "Arial"};
            var supplementalFont37 = new A.SupplementalFont() {Script = "Thai", Typeface = "Cordia New"};
            var supplementalFont38 = new A.SupplementalFont() {Script = "Ethi", Typeface = "Nyala"};
            var supplementalFont39 = new A.SupplementalFont() {Script = "Beng", Typeface = "Vrinda"};
            var supplementalFont40 = new A.SupplementalFont() {Script = "Gujr", Typeface = "Shruti"};
            var supplementalFont41 = new A.SupplementalFont() {Script = "Khmr", Typeface = "DaunPenh"};
            var supplementalFont42 = new A.SupplementalFont() {Script = "Knda", Typeface = "Tunga"};
            var supplementalFont43 = new A.SupplementalFont() {Script = "Guru", Typeface = "Raavi"};
            var supplementalFont44 = new A.SupplementalFont() {Script = "Cans", Typeface = "Euphemia"};
            var supplementalFont45 = new A.SupplementalFont() {Script = "Cher", Typeface = "Plantagenet Cherokee"};
            var supplementalFont46 = new A.SupplementalFont() {Script = "Yiii", Typeface = "Microsoft Yi Baiti"};
            var supplementalFont47 = new A.SupplementalFont() {Script = "Tibt", Typeface = "Microsoft Himalaya"};
            var supplementalFont48 = new A.SupplementalFont() {Script = "Thaa", Typeface = "MV Boli"};
            var supplementalFont49 = new A.SupplementalFont() {Script = "Deva", Typeface = "Mangal"};
            var supplementalFont50 = new A.SupplementalFont() {Script = "Telu", Typeface = "Gautami"};
            var supplementalFont51 = new A.SupplementalFont() {Script = "Taml", Typeface = "Latha"};
            var supplementalFont52 = new A.SupplementalFont() {Script = "Syrc", Typeface = "Estrangelo Edessa"};
            var supplementalFont53 = new A.SupplementalFont() {Script = "Orya", Typeface = "Kalinga"};
            var supplementalFont54 = new A.SupplementalFont() {Script = "Mlym", Typeface = "Kartika"};
            var supplementalFont55 = new A.SupplementalFont() {Script = "Laoo", Typeface = "DokChampa"};
            var supplementalFont56 = new A.SupplementalFont() {Script = "Sinh", Typeface = "Iskoola Pota"};
            var supplementalFont57 = new A.SupplementalFont() {Script = "Mong", Typeface = "Mongolian Baiti"};
            var supplementalFont58 = new A.SupplementalFont() {Script = "Viet", Typeface = "Arial"};
            var supplementalFont59 = new A.SupplementalFont() {Script = "Uigh", Typeface = "Microsoft Uighur"};
            var supplementalFont60 = new A.SupplementalFont() {Script = "Geor", Typeface = "Sylfaen"};

            minorFont1.Append(latinFont30);
            minorFont1.Append(eastAsianFont30);
            minorFont1.Append(complexScriptFont30);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            var formatScheme1 = new A.FormatScheme() {Name = "Office"};

            var fillStyleList1 = new A.FillStyleList();

            var solidFill41 = new A.SolidFill();
            var schemeColor42 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};

            solidFill41.Append(schemeColor42);

            var gradientFill1 = new A.GradientFill() {RotateWithShape = true};

            var gradientStopList1 = new A.GradientStopList();

            var gradientStop1 = new A.GradientStop() {Position = 0};

            var schemeColor43 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var tint13 = new A.Tint() {Val = 50000};
            var saturationModulation1 = new A.SaturationModulation() {Val = 300000};

            schemeColor43.Append(tint13);
            schemeColor43.Append(saturationModulation1);

            gradientStop1.Append(schemeColor43);

            var gradientStop2 = new A.GradientStop() {Position = 35000};

            var schemeColor44 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var tint14 = new A.Tint() {Val = 37000};
            var saturationModulation2 = new A.SaturationModulation() {Val = 300000};

            schemeColor44.Append(tint14);
            schemeColor44.Append(saturationModulation2);

            gradientStop2.Append(schemeColor44);

            var gradientStop3 = new A.GradientStop() {Position = 100000};

            var schemeColor45 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var tint15 = new A.Tint() {Val = 15000};
            var saturationModulation3 = new A.SaturationModulation() {Val = 350000};

            schemeColor45.Append(tint15);
            schemeColor45.Append(saturationModulation3);

            gradientStop3.Append(schemeColor45);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            var linearGradientFill1 = new A.LinearGradientFill() {Angle = 16200000, Scaled = true};

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            var gradientFill2 = new A.GradientFill() {RotateWithShape = true};

            var gradientStopList2 = new A.GradientStopList();

            var gradientStop4 = new A.GradientStop() {Position = 0};

            var schemeColor46 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var shade1 = new A.Shade() {Val = 51000};
            var saturationModulation4 = new A.SaturationModulation() {Val = 130000};

            schemeColor46.Append(shade1);
            schemeColor46.Append(saturationModulation4);

            gradientStop4.Append(schemeColor46);

            var gradientStop5 = new A.GradientStop() {Position = 80000};

            var schemeColor47 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var shade2 = new A.Shade() {Val = 93000};
            var saturationModulation5 = new A.SaturationModulation() {Val = 130000};

            schemeColor47.Append(shade2);
            schemeColor47.Append(saturationModulation5);

            gradientStop5.Append(schemeColor47);

            var gradientStop6 = new A.GradientStop() {Position = 100000};

            var schemeColor48 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var shade3 = new A.Shade() {Val = 94000};
            var saturationModulation6 = new A.SaturationModulation() {Val = 135000};

            schemeColor48.Append(shade3);
            schemeColor48.Append(saturationModulation6);

            gradientStop6.Append(schemeColor48);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            var linearGradientFill2 = new A.LinearGradientFill() {Angle = 16200000, Scaled = false};

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill41);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            var lineStyleList1 = new A.LineStyleList();

            var outline1 = new A.Outline() {
                Width = 9525,
                CapType = A.LineCapValues.Flat,
                CompoundLineType = A.CompoundLineValues.Single,
                Alignment = A.PenAlignmentValues.Center
            };

            var solidFill42 = new A.SolidFill();

            var schemeColor49 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var shade4 = new A.Shade() {Val = 95000};
            var saturationModulation7 = new A.SaturationModulation() {Val = 105000};

            schemeColor49.Append(shade4);
            schemeColor49.Append(saturationModulation7);

            solidFill42.Append(schemeColor49);
            var presetDash1 = new A.PresetDash() {Val = A.PresetLineDashValues.Solid};

            outline1.Append(solidFill42);
            outline1.Append(presetDash1);

            var outline2 = new A.Outline() {
                Width = 25400,
                CapType = A.LineCapValues.Flat,
                CompoundLineType = A.CompoundLineValues.Single,
                Alignment = A.PenAlignmentValues.Center
            };

            var solidFill43 = new A.SolidFill();
            var schemeColor50 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};

            solidFill43.Append(schemeColor50);
            var presetDash2 = new A.PresetDash() {Val = A.PresetLineDashValues.Solid};

            outline2.Append(solidFill43);
            outline2.Append(presetDash2);

            var outline3 = new A.Outline() {
                Width = 38100,
                CapType = A.LineCapValues.Flat,
                CompoundLineType = A.CompoundLineValues.Single,
                Alignment = A.PenAlignmentValues.Center
            };

            var solidFill44 = new A.SolidFill();
            var schemeColor51 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};

            solidFill44.Append(schemeColor51);
            var presetDash3 = new A.PresetDash() {Val = A.PresetLineDashValues.Solid};

            outline3.Append(solidFill44);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            var effectStyleList1 = new A.EffectStyleList();

            var effectStyle1 = new A.EffectStyle();

            var effectList1 = new A.EffectList();

            var outerShadow1 = new A.OuterShadow() {
                BlurRadius = 40000L,
                Distance = 20000L,
                Direction = 5400000,
                RotateWithShape = false
            };

            var rgbColorModelHex11 = new A.RgbColorModelHex() {Val = "000000"};
            var alpha1 = new A.Alpha() {Val = 38000};

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            var effectStyle2 = new A.EffectStyle();

            var effectList2 = new A.EffectList();

            var outerShadow2 = new A.OuterShadow() {
                BlurRadius = 40000L,
                Distance = 23000L,
                Direction = 5400000,
                RotateWithShape = false
            };

            var rgbColorModelHex12 = new A.RgbColorModelHex() {Val = "000000"};
            var alpha2 = new A.Alpha() {Val = 35000};

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            var effectStyle3 = new A.EffectStyle();

            var effectList3 = new A.EffectList();

            var outerShadow3 = new A.OuterShadow() {
                BlurRadius = 40000L,
                Distance = 23000L,
                Direction = 5400000,
                RotateWithShape = false
            };

            var rgbColorModelHex13 = new A.RgbColorModelHex() {Val = "000000"};
            var alpha3 = new A.Alpha() {Val = 35000};

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            var scene3DType1 = new A.Scene3DType();

            var camera1 = new A.Camera() {Preset = A.PresetCameraValues.OrthographicFront};
            var rotation1 = new A.Rotation() {Latitude = 0, Longitude = 0, Revolution = 0};

            camera1.Append(rotation1);

            var lightRig1 = new A.LightRig() {
                Rig = A.LightRigValues.ThreePoints,
                Direction = A.LightRigDirectionValues.Top
            };
            var rotation2 = new A.Rotation() {Latitude = 0, Longitude = 0, Revolution = 1200000};

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            var shape3DType1 = new A.Shape3DType();
            var bevelTop1 = new A.BevelTop() {Width = 63500L, Height = 25400L};

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            var backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            var solidFill45 = new A.SolidFill();
            var schemeColor52 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};

            solidFill45.Append(schemeColor52);

            var gradientFill3 = new A.GradientFill() {RotateWithShape = true};

            var gradientStopList3 = new A.GradientStopList();

            var gradientStop7 = new A.GradientStop() {Position = 0};

            var schemeColor53 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var tint16 = new A.Tint() {Val = 40000};
            var saturationModulation8 = new A.SaturationModulation() {Val = 350000};

            schemeColor53.Append(tint16);
            schemeColor53.Append(saturationModulation8);

            gradientStop7.Append(schemeColor53);

            var gradientStop8 = new A.GradientStop() {Position = 40000};

            var schemeColor54 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var tint17 = new A.Tint() {Val = 45000};
            var shade5 = new A.Shade() {Val = 99000};
            var saturationModulation9 = new A.SaturationModulation() {Val = 350000};

            schemeColor54.Append(tint17);
            schemeColor54.Append(shade5);
            schemeColor54.Append(saturationModulation9);

            gradientStop8.Append(schemeColor54);

            var gradientStop9 = new A.GradientStop() {Position = 100000};

            var schemeColor55 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var shade6 = new A.Shade() {Val = 20000};
            var saturationModulation10 = new A.SaturationModulation() {Val = 255000};

            schemeColor55.Append(shade6);
            schemeColor55.Append(saturationModulation10);

            gradientStop9.Append(schemeColor55);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            var pathGradientFill1 = new A.PathGradientFill() {Path = A.PathShadeValues.Circle};
            var fillToRectangle1 = new A.FillToRectangle() {Left = 50000, Top = -80000, Right = 50000, Bottom = 180000};

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            var gradientFill4 = new A.GradientFill() {RotateWithShape = true};

            var gradientStopList4 = new A.GradientStopList();

            var gradientStop10 = new A.GradientStop() {Position = 0};

            var schemeColor56 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var tint18 = new A.Tint() {Val = 80000};
            var saturationModulation11 = new A.SaturationModulation() {Val = 300000};

            schemeColor56.Append(tint18);
            schemeColor56.Append(saturationModulation11);

            gradientStop10.Append(schemeColor56);

            var gradientStop11 = new A.GradientStop() {Position = 100000};

            var schemeColor57 = new A.SchemeColor() {Val = A.SchemeColorValues.PhColor};
            var shade7 = new A.Shade() {Val = 30000};
            var saturationModulation12 = new A.SaturationModulation() {Val = 200000};

            schemeColor57.Append(shade7);
            schemeColor57.Append(saturationModulation12);

            gradientStop11.Append(schemeColor57);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            var pathGradientFill2 = new A.PathGradientFill() {Path = A.PathShadeValues.Circle};
            var fillToRectangle2 = new A.FillToRectangle() {Left = 50000, Top = 50000, Right = 50000, Bottom = 50000};

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill45);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            var objectDefaults1 = new A.ObjectDefaults();
            var extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        // Generates content of slideLayoutPart5.
        private static void GenerateSlideLayoutPart5Content(SlideLayoutPart slideLayoutPart5)
        {
            var slideLayout5 = new SlideLayout() {Type = SlideLayoutValues.Object, Preserve = true};
            slideLayout5.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout5.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout5.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData7 = new CommonSlideData() {Name = "Title and Content"};

            var shapeTree7 = new ShapeTree();

            var nonVisualGroupShapeProperties7 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties31 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties7 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties31 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties7.Append(nonVisualDrawingProperties31);
            nonVisualGroupShapeProperties7.Append(nonVisualGroupShapeDrawingProperties7);
            nonVisualGroupShapeProperties7.Append(applicationNonVisualDrawingProperties31);

            var groupShapeProperties7 = new GroupShapeProperties();

            var transformGroup7 = new A.TransformGroup();
            var offset17 = new A.Offset() {X = 0L, Y = 0L};
            var extents17 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset7 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents7 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup7.Append(offset17);
            transformGroup7.Append(extents17);
            transformGroup7.Append(childOffset7);
            transformGroup7.Append(childExtents7);

            groupShapeProperties7.Append(transformGroup7);

            var shape25 = new Shape();

            var nonVisualShapeProperties25 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties32 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 2U,
                Name = "Title 1"
            };

            var nonVisualShapeDrawingProperties25 = new NonVisualShapeDrawingProperties();
            var shapeLocks25 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties25.Append(shapeLocks25);

            var applicationNonVisualDrawingProperties32 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape25 = new PlaceholderShape() {Type = PlaceholderValues.Title};

            applicationNonVisualDrawingProperties32.Append(placeholderShape25);

            nonVisualShapeProperties25.Append(nonVisualDrawingProperties32);
            nonVisualShapeProperties25.Append(nonVisualShapeDrawingProperties25);
            nonVisualShapeProperties25.Append(applicationNonVisualDrawingProperties32);
            var shapeProperties25 = new ShapeProperties();

            var textBody25 = new TextBody();
            var bodyProperties25 = new A.BodyProperties();
            var listStyle25 = new A.ListStyle();

            var paragraph33 = new A.Paragraph();

            var run17 = new A.Run();

            var runProperties27 = new A.RunProperties() {Language = "en-US"};
            runProperties27.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text27 = new A.Text();
            text27.Text = "Click to edit Master title style";

            run17.Append(runProperties27);
            run17.Append(text27);
            var endParagraphRunProperties23 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph33.Append(run17);
            paragraph33.Append(endParagraphRunProperties23);

            textBody25.Append(bodyProperties25);
            textBody25.Append(listStyle25);
            textBody25.Append(paragraph33);

            shape25.Append(nonVisualShapeProperties25);
            shape25.Append(shapeProperties25);
            shape25.Append(textBody25);

            var shape26 = new Shape();

            var nonVisualShapeProperties26 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties33 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Content Placeholder 2"
            };

            var nonVisualShapeDrawingProperties26 = new NonVisualShapeDrawingProperties();
            var shapeLocks26 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties26.Append(shapeLocks26);

            var applicationNonVisualDrawingProperties33 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape26 = new PlaceholderShape() {Index = (UInt32Value) 1U};

            applicationNonVisualDrawingProperties33.Append(placeholderShape26);

            nonVisualShapeProperties26.Append(nonVisualDrawingProperties33);
            nonVisualShapeProperties26.Append(nonVisualShapeDrawingProperties26);
            nonVisualShapeProperties26.Append(applicationNonVisualDrawingProperties33);
            var shapeProperties26 = new ShapeProperties();

            var textBody26 = new TextBody();
            var bodyProperties26 = new A.BodyProperties();
            var listStyle26 = new A.ListStyle();

            var paragraph34 = new A.Paragraph();
            var paragraphProperties13 = new A.ParagraphProperties() {Level = 0};

            var run18 = new A.Run();

            var runProperties28 = new A.RunProperties() {Language = "en-US"};
            runProperties28.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text28 = new A.Text();
            text28.Text = "Click to edit Master text styles";

            run18.Append(runProperties28);
            run18.Append(text28);

            paragraph34.Append(paragraphProperties13);
            paragraph34.Append(run18);

            var paragraph35 = new A.Paragraph();
            var paragraphProperties14 = new A.ParagraphProperties() {Level = 1};

            var run19 = new A.Run();

            var runProperties29 = new A.RunProperties() {Language = "en-US"};
            runProperties29.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text29 = new A.Text();
            text29.Text = "Second level";

            run19.Append(runProperties29);
            run19.Append(text29);

            paragraph35.Append(paragraphProperties14);
            paragraph35.Append(run19);

            var paragraph36 = new A.Paragraph();
            var paragraphProperties15 = new A.ParagraphProperties() {Level = 2};

            var run20 = new A.Run();

            var runProperties30 = new A.RunProperties() {Language = "en-US"};
            runProperties30.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text30 = new A.Text();
            text30.Text = "Third level";

            run20.Append(runProperties30);
            run20.Append(text30);

            paragraph36.Append(paragraphProperties15);
            paragraph36.Append(run20);

            var paragraph37 = new A.Paragraph();
            var paragraphProperties16 = new A.ParagraphProperties() {Level = 3};

            var run21 = new A.Run();

            var runProperties31 = new A.RunProperties() {Language = "en-US"};
            runProperties31.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text31 = new A.Text();
            text31.Text = "Fourth level";

            run21.Append(runProperties31);
            run21.Append(text31);

            paragraph37.Append(paragraphProperties16);
            paragraph37.Append(run21);

            var paragraph38 = new A.Paragraph();
            var paragraphProperties17 = new A.ParagraphProperties() {Level = 4};

            var run22 = new A.Run();

            var runProperties32 = new A.RunProperties() {Language = "en-US"};
            runProperties32.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text32 = new A.Text();
            text32.Text = "Fifth level";

            run22.Append(runProperties32);
            run22.Append(text32);
            var endParagraphRunProperties24 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph38.Append(paragraphProperties17);
            paragraph38.Append(run22);
            paragraph38.Append(endParagraphRunProperties24);

            textBody26.Append(bodyProperties26);
            textBody26.Append(listStyle26);
            textBody26.Append(paragraph34);
            textBody26.Append(paragraph35);
            textBody26.Append(paragraph36);
            textBody26.Append(paragraph37);
            textBody26.Append(paragraph38);

            shape26.Append(nonVisualShapeProperties26);
            shape26.Append(shapeProperties26);
            shape26.Append(textBody26);

            var shape27 = new Shape();

            var nonVisualShapeProperties27 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties34 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Date Placeholder 3"
            };

            var nonVisualShapeDrawingProperties27 = new NonVisualShapeDrawingProperties();
            var shapeLocks27 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties27.Append(shapeLocks27);

            var applicationNonVisualDrawingProperties34 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape27 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 10U
            };

            applicationNonVisualDrawingProperties34.Append(placeholderShape27);

            nonVisualShapeProperties27.Append(nonVisualDrawingProperties34);
            nonVisualShapeProperties27.Append(nonVisualShapeDrawingProperties27);
            nonVisualShapeProperties27.Append(applicationNonVisualDrawingProperties34);
            var shapeProperties27 = new ShapeProperties();

            var textBody27 = new TextBody();
            var bodyProperties27 = new A.BodyProperties();
            var listStyle27 = new A.ListStyle();

            var paragraph39 = new A.Paragraph();

            var field11 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties33 = new A.RunProperties() {Language = "en-GB"};
            runProperties33.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text33 = new A.Text();
            text33.Text = "04/06/2014";

            field11.Append(runProperties33);
            field11.Append(text33);
            var endParagraphRunProperties25 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph39.Append(field11);
            paragraph39.Append(endParagraphRunProperties25);

            textBody27.Append(bodyProperties27);
            textBody27.Append(listStyle27);
            textBody27.Append(paragraph39);

            shape27.Append(nonVisualShapeProperties27);
            shape27.Append(shapeProperties27);
            shape27.Append(textBody27);

            var shape28 = new Shape();

            var nonVisualShapeProperties28 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties35 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 5U,
                Name = "Footer Placeholder 4"
            };

            var nonVisualShapeDrawingProperties28 = new NonVisualShapeDrawingProperties();
            var shapeLocks28 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties28.Append(shapeLocks28);

            var applicationNonVisualDrawingProperties35 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape28 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 11U
            };

            applicationNonVisualDrawingProperties35.Append(placeholderShape28);

            nonVisualShapeProperties28.Append(nonVisualDrawingProperties35);
            nonVisualShapeProperties28.Append(nonVisualShapeDrawingProperties28);
            nonVisualShapeProperties28.Append(applicationNonVisualDrawingProperties35);
            var shapeProperties28 = new ShapeProperties();

            var textBody28 = new TextBody();
            var bodyProperties28 = new A.BodyProperties();
            var listStyle28 = new A.ListStyle();

            var paragraph40 = new A.Paragraph();
            var endParagraphRunProperties26 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph40.Append(endParagraphRunProperties26);

            textBody28.Append(bodyProperties28);
            textBody28.Append(listStyle28);
            textBody28.Append(paragraph40);

            shape28.Append(nonVisualShapeProperties28);
            shape28.Append(shapeProperties28);
            shape28.Append(textBody28);

            var shape29 = new Shape();

            var nonVisualShapeProperties29 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties36 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 6U,
                Name = "Slide Number Placeholder 5"
            };

            var nonVisualShapeDrawingProperties29 = new NonVisualShapeDrawingProperties();
            var shapeLocks29 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties29.Append(shapeLocks29);

            var applicationNonVisualDrawingProperties36 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape29 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 12U
            };

            applicationNonVisualDrawingProperties36.Append(placeholderShape29);

            nonVisualShapeProperties29.Append(nonVisualDrawingProperties36);
            nonVisualShapeProperties29.Append(nonVisualShapeDrawingProperties29);
            nonVisualShapeProperties29.Append(applicationNonVisualDrawingProperties36);
            var shapeProperties29 = new ShapeProperties();

            var textBody29 = new TextBody();
            var bodyProperties29 = new A.BodyProperties();
            var listStyle29 = new A.ListStyle();

            var paragraph41 = new A.Paragraph();

            var field12 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties34 = new A.RunProperties() {Language = "en-GB"};
            runProperties34.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text34 = new A.Text();
            text34.Text = "‹#›";

            field12.Append(runProperties34);
            field12.Append(text34);
            var endParagraphRunProperties27 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph41.Append(field12);
            paragraph41.Append(endParagraphRunProperties27);

            textBody29.Append(bodyProperties29);
            textBody29.Append(listStyle29);
            textBody29.Append(paragraph41);

            shape29.Append(nonVisualShapeProperties29);
            shape29.Append(shapeProperties29);
            shape29.Append(textBody29);

            shapeTree7.Append(nonVisualGroupShapeProperties7);
            shapeTree7.Append(groupShapeProperties7);
            shapeTree7.Append(shape25);
            shapeTree7.Append(shape26);
            shapeTree7.Append(shape27);
            shapeTree7.Append(shape28);
            shapeTree7.Append(shape29);

            var commonSlideDataExtensionList7 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension7 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId7 = new P14.CreationId() {Val = (UInt32Value) 2090794633U};
            creationId7.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension7.Append(creationId7);

            commonSlideDataExtensionList7.Append(commonSlideDataExtension7);

            commonSlideData7.Append(shapeTree7);
            commonSlideData7.Append(commonSlideDataExtensionList7);

            var colorMapOverride6 = new ColorMapOverride();
            var masterColorMapping6 = new A.MasterColorMapping();

            colorMapOverride6.Append(masterColorMapping6);

            slideLayout5.Append(commonSlideData7);
            slideLayout5.Append(colorMapOverride6);

            slideLayoutPart5.SlideLayout = slideLayout5;
        }

        // Generates content of slideLayoutPart6.
        private static void GenerateSlideLayoutPart6Content(SlideLayoutPart slideLayoutPart6)
        {
            var slideLayout6 = new SlideLayout() {Type = SlideLayoutValues.Title, Preserve = true};
            slideLayout6.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout6.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout6.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData8 = new CommonSlideData() {Name = "Title Slide"};

            var shapeTree8 = new ShapeTree();

            var nonVisualGroupShapeProperties8 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties37 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties8 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties37 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties8.Append(nonVisualDrawingProperties37);
            nonVisualGroupShapeProperties8.Append(nonVisualGroupShapeDrawingProperties8);
            nonVisualGroupShapeProperties8.Append(applicationNonVisualDrawingProperties37);

            var groupShapeProperties8 = new GroupShapeProperties();

            var transformGroup8 = new A.TransformGroup();
            var offset18 = new A.Offset() {X = 0L, Y = 0L};
            var extents18 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset8 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents8 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup8.Append(offset18);
            transformGroup8.Append(extents18);
            transformGroup8.Append(childOffset8);
            transformGroup8.Append(childExtents8);

            groupShapeProperties8.Append(transformGroup8);

            var shape30 = new Shape();

            var nonVisualShapeProperties30 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties38 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 2U,
                Name = "Title 1"
            };

            var nonVisualShapeDrawingProperties30 = new NonVisualShapeDrawingProperties();
            var shapeLocks30 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties30.Append(shapeLocks30);

            var applicationNonVisualDrawingProperties38 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape30 = new PlaceholderShape() {Type = PlaceholderValues.CenteredTitle};

            applicationNonVisualDrawingProperties38.Append(placeholderShape30);

            nonVisualShapeProperties30.Append(nonVisualDrawingProperties38);
            nonVisualShapeProperties30.Append(nonVisualShapeDrawingProperties30);
            nonVisualShapeProperties30.Append(applicationNonVisualDrawingProperties38);

            var shapeProperties30 = new ShapeProperties();

            var transform2D11 = new A.Transform2D();
            var offset19 = new A.Offset() {X = 685800L, Y = 2130425L};
            var extents19 = new A.Extents() {Cx = 7772400L, Cy = 1470025L};

            transform2D11.Append(offset19);
            transform2D11.Append(extents19);

            shapeProperties30.Append(transform2D11);

            var textBody30 = new TextBody();
            var bodyProperties30 = new A.BodyProperties();
            var listStyle30 = new A.ListStyle();

            var paragraph42 = new A.Paragraph();

            var run23 = new A.Run();

            var runProperties35 = new A.RunProperties() {Language = "en-US"};
            runProperties35.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text35 = new A.Text();
            text35.Text = "Click to edit Master title style";

            run23.Append(runProperties35);
            run23.Append(text35);
            var endParagraphRunProperties28 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph42.Append(run23);
            paragraph42.Append(endParagraphRunProperties28);

            textBody30.Append(bodyProperties30);
            textBody30.Append(listStyle30);
            textBody30.Append(paragraph42);

            shape30.Append(nonVisualShapeProperties30);
            shape30.Append(shapeProperties30);
            shape30.Append(textBody30);

            var shape31 = new Shape();

            var nonVisualShapeProperties31 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties39 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Subtitle 2"
            };

            var nonVisualShapeDrawingProperties31 = new NonVisualShapeDrawingProperties();
            var shapeLocks31 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties31.Append(shapeLocks31);

            var applicationNonVisualDrawingProperties39 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape31 = new PlaceholderShape() {
                Type = PlaceholderValues.SubTitle,
                Index = (UInt32Value) 1U
            };

            applicationNonVisualDrawingProperties39.Append(placeholderShape31);

            nonVisualShapeProperties31.Append(nonVisualDrawingProperties39);
            nonVisualShapeProperties31.Append(nonVisualShapeDrawingProperties31);
            nonVisualShapeProperties31.Append(applicationNonVisualDrawingProperties39);

            var shapeProperties31 = new ShapeProperties();

            var transform2D12 = new A.Transform2D();
            var offset20 = new A.Offset() {X = 1371600L, Y = 3886200L};
            var extents20 = new A.Extents() {Cx = 6400800L, Cy = 1752600L};

            transform2D12.Append(offset20);
            transform2D12.Append(extents20);

            shapeProperties31.Append(transform2D12);

            var textBody31 = new TextBody();
            var bodyProperties31 = new A.BodyProperties();

            var listStyle31 = new A.ListStyle();

            var level1ParagraphProperties13 = new A.Level1ParagraphProperties() {
                LeftMargin = 0,
                Indent = 0,
                Alignment = A.TextAlignmentTypeValues.Center
            };
            var noBullet20 = new A.NoBullet();

            var defaultRunProperties63 = new A.DefaultRunProperties();

            var solidFill46 = new A.SolidFill();

            var schemeColor58 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint19 = new A.Tint() {Val = 75000};

            schemeColor58.Append(tint19);

            solidFill46.Append(schemeColor58);

            defaultRunProperties63.Append(solidFill46);

            level1ParagraphProperties13.Append(noBullet20);
            level1ParagraphProperties13.Append(defaultRunProperties63);

            var level2ParagraphProperties7 = new A.Level2ParagraphProperties() {
                LeftMargin = 457200,
                Indent = 0,
                Alignment = A.TextAlignmentTypeValues.Center
            };
            var noBullet21 = new A.NoBullet();

            var defaultRunProperties64 = new A.DefaultRunProperties();

            var solidFill47 = new A.SolidFill();

            var schemeColor59 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint20 = new A.Tint() {Val = 75000};

            schemeColor59.Append(tint20);

            solidFill47.Append(schemeColor59);

            defaultRunProperties64.Append(solidFill47);

            level2ParagraphProperties7.Append(noBullet21);
            level2ParagraphProperties7.Append(defaultRunProperties64);

            var level3ParagraphProperties7 = new A.Level3ParagraphProperties() {
                LeftMargin = 914400,
                Indent = 0,
                Alignment = A.TextAlignmentTypeValues.Center
            };
            var noBullet22 = new A.NoBullet();

            var defaultRunProperties65 = new A.DefaultRunProperties();

            var solidFill48 = new A.SolidFill();

            var schemeColor60 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint21 = new A.Tint() {Val = 75000};

            schemeColor60.Append(tint21);

            solidFill48.Append(schemeColor60);

            defaultRunProperties65.Append(solidFill48);

            level3ParagraphProperties7.Append(noBullet22);
            level3ParagraphProperties7.Append(defaultRunProperties65);

            var level4ParagraphProperties7 = new A.Level4ParagraphProperties() {
                LeftMargin = 1371600,
                Indent = 0,
                Alignment = A.TextAlignmentTypeValues.Center
            };
            var noBullet23 = new A.NoBullet();

            var defaultRunProperties66 = new A.DefaultRunProperties();

            var solidFill49 = new A.SolidFill();

            var schemeColor61 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint22 = new A.Tint() {Val = 75000};

            schemeColor61.Append(tint22);

            solidFill49.Append(schemeColor61);

            defaultRunProperties66.Append(solidFill49);

            level4ParagraphProperties7.Append(noBullet23);
            level4ParagraphProperties7.Append(defaultRunProperties66);

            var level5ParagraphProperties7 = new A.Level5ParagraphProperties() {
                LeftMargin = 1828800,
                Indent = 0,
                Alignment = A.TextAlignmentTypeValues.Center
            };
            var noBullet24 = new A.NoBullet();

            var defaultRunProperties67 = new A.DefaultRunProperties();

            var solidFill50 = new A.SolidFill();

            var schemeColor62 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint23 = new A.Tint() {Val = 75000};

            schemeColor62.Append(tint23);

            solidFill50.Append(schemeColor62);

            defaultRunProperties67.Append(solidFill50);

            level5ParagraphProperties7.Append(noBullet24);
            level5ParagraphProperties7.Append(defaultRunProperties67);

            var level6ParagraphProperties7 = new A.Level6ParagraphProperties() {
                LeftMargin = 2286000,
                Indent = 0,
                Alignment = A.TextAlignmentTypeValues.Center
            };
            var noBullet25 = new A.NoBullet();

            var defaultRunProperties68 = new A.DefaultRunProperties();

            var solidFill51 = new A.SolidFill();

            var schemeColor63 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint24 = new A.Tint() {Val = 75000};

            schemeColor63.Append(tint24);

            solidFill51.Append(schemeColor63);

            defaultRunProperties68.Append(solidFill51);

            level6ParagraphProperties7.Append(noBullet25);
            level6ParagraphProperties7.Append(defaultRunProperties68);

            var level7ParagraphProperties7 = new A.Level7ParagraphProperties() {
                LeftMargin = 2743200,
                Indent = 0,
                Alignment = A.TextAlignmentTypeValues.Center
            };
            var noBullet26 = new A.NoBullet();

            var defaultRunProperties69 = new A.DefaultRunProperties();

            var solidFill52 = new A.SolidFill();

            var schemeColor64 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint25 = new A.Tint() {Val = 75000};

            schemeColor64.Append(tint25);

            solidFill52.Append(schemeColor64);

            defaultRunProperties69.Append(solidFill52);

            level7ParagraphProperties7.Append(noBullet26);
            level7ParagraphProperties7.Append(defaultRunProperties69);

            var level8ParagraphProperties7 = new A.Level8ParagraphProperties() {
                LeftMargin = 3200400,
                Indent = 0,
                Alignment = A.TextAlignmentTypeValues.Center
            };
            var noBullet27 = new A.NoBullet();

            var defaultRunProperties70 = new A.DefaultRunProperties();

            var solidFill53 = new A.SolidFill();

            var schemeColor65 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint26 = new A.Tint() {Val = 75000};

            schemeColor65.Append(tint26);

            solidFill53.Append(schemeColor65);

            defaultRunProperties70.Append(solidFill53);

            level8ParagraphProperties7.Append(noBullet27);
            level8ParagraphProperties7.Append(defaultRunProperties70);

            var level9ParagraphProperties7 = new A.Level9ParagraphProperties() {
                LeftMargin = 3657600,
                Indent = 0,
                Alignment = A.TextAlignmentTypeValues.Center
            };
            var noBullet28 = new A.NoBullet();

            var defaultRunProperties71 = new A.DefaultRunProperties();

            var solidFill54 = new A.SolidFill();

            var schemeColor66 = new A.SchemeColor() {Val = A.SchemeColorValues.Text1};
            var tint27 = new A.Tint() {Val = 75000};

            schemeColor66.Append(tint27);

            solidFill54.Append(schemeColor66);

            defaultRunProperties71.Append(solidFill54);

            level9ParagraphProperties7.Append(noBullet28);
            level9ParagraphProperties7.Append(defaultRunProperties71);

            listStyle31.Append(level1ParagraphProperties13);
            listStyle31.Append(level2ParagraphProperties7);
            listStyle31.Append(level3ParagraphProperties7);
            listStyle31.Append(level4ParagraphProperties7);
            listStyle31.Append(level5ParagraphProperties7);
            listStyle31.Append(level6ParagraphProperties7);
            listStyle31.Append(level7ParagraphProperties7);
            listStyle31.Append(level8ParagraphProperties7);
            listStyle31.Append(level9ParagraphProperties7);

            var paragraph43 = new A.Paragraph();

            var run24 = new A.Run();

            var runProperties36 = new A.RunProperties() {Language = "en-US"};
            runProperties36.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text36 = new A.Text();
            text36.Text = "Click to edit Master subtitle style";

            run24.Append(runProperties36);
            run24.Append(text36);
            var endParagraphRunProperties29 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph43.Append(run24);
            paragraph43.Append(endParagraphRunProperties29);

            textBody31.Append(bodyProperties31);
            textBody31.Append(listStyle31);
            textBody31.Append(paragraph43);

            shape31.Append(nonVisualShapeProperties31);
            shape31.Append(shapeProperties31);
            shape31.Append(textBody31);

            var shape32 = new Shape();

            var nonVisualShapeProperties32 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties40 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Date Placeholder 3"
            };

            var nonVisualShapeDrawingProperties32 = new NonVisualShapeDrawingProperties();
            var shapeLocks32 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties32.Append(shapeLocks32);

            var applicationNonVisualDrawingProperties40 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape32 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 10U
            };

            applicationNonVisualDrawingProperties40.Append(placeholderShape32);

            nonVisualShapeProperties32.Append(nonVisualDrawingProperties40);
            nonVisualShapeProperties32.Append(nonVisualShapeDrawingProperties32);
            nonVisualShapeProperties32.Append(applicationNonVisualDrawingProperties40);
            var shapeProperties32 = new ShapeProperties();

            var textBody32 = new TextBody();
            var bodyProperties32 = new A.BodyProperties();
            var listStyle32 = new A.ListStyle();

            var paragraph44 = new A.Paragraph();

            var field13 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties37 = new A.RunProperties() {Language = "en-GB"};
            runProperties37.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text37 = new A.Text();
            text37.Text = "04/06/2014";

            field13.Append(runProperties37);
            field13.Append(text37);
            var endParagraphRunProperties30 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph44.Append(field13);
            paragraph44.Append(endParagraphRunProperties30);

            textBody32.Append(bodyProperties32);
            textBody32.Append(listStyle32);
            textBody32.Append(paragraph44);

            shape32.Append(nonVisualShapeProperties32);
            shape32.Append(shapeProperties32);
            shape32.Append(textBody32);

            var shape33 = new Shape();

            var nonVisualShapeProperties33 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties41 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 5U,
                Name = "Footer Placeholder 4"
            };

            var nonVisualShapeDrawingProperties33 = new NonVisualShapeDrawingProperties();
            var shapeLocks33 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties33.Append(shapeLocks33);

            var applicationNonVisualDrawingProperties41 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape33 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 11U
            };

            applicationNonVisualDrawingProperties41.Append(placeholderShape33);

            nonVisualShapeProperties33.Append(nonVisualDrawingProperties41);
            nonVisualShapeProperties33.Append(nonVisualShapeDrawingProperties33);
            nonVisualShapeProperties33.Append(applicationNonVisualDrawingProperties41);
            var shapeProperties33 = new ShapeProperties();

            var textBody33 = new TextBody();
            var bodyProperties33 = new A.BodyProperties();
            var listStyle33 = new A.ListStyle();

            var paragraph45 = new A.Paragraph();
            var endParagraphRunProperties31 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph45.Append(endParagraphRunProperties31);

            textBody33.Append(bodyProperties33);
            textBody33.Append(listStyle33);
            textBody33.Append(paragraph45);

            shape33.Append(nonVisualShapeProperties33);
            shape33.Append(shapeProperties33);
            shape33.Append(textBody33);

            var shape34 = new Shape();

            var nonVisualShapeProperties34 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties42 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 6U,
                Name = "Slide Number Placeholder 5"
            };

            var nonVisualShapeDrawingProperties34 = new NonVisualShapeDrawingProperties();
            var shapeLocks34 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties34.Append(shapeLocks34);

            var applicationNonVisualDrawingProperties42 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape34 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 12U
            };

            applicationNonVisualDrawingProperties42.Append(placeholderShape34);

            nonVisualShapeProperties34.Append(nonVisualDrawingProperties42);
            nonVisualShapeProperties34.Append(nonVisualShapeDrawingProperties34);
            nonVisualShapeProperties34.Append(applicationNonVisualDrawingProperties42);
            var shapeProperties34 = new ShapeProperties();

            var textBody34 = new TextBody();
            var bodyProperties34 = new A.BodyProperties();
            var listStyle34 = new A.ListStyle();

            var paragraph46 = new A.Paragraph();

            var field14 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties38 = new A.RunProperties() {Language = "en-GB"};
            runProperties38.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text38 = new A.Text();
            text38.Text = "‹#›";

            field14.Append(runProperties38);
            field14.Append(text38);
            var endParagraphRunProperties32 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph46.Append(field14);
            paragraph46.Append(endParagraphRunProperties32);

            textBody34.Append(bodyProperties34);
            textBody34.Append(listStyle34);
            textBody34.Append(paragraph46);

            shape34.Append(nonVisualShapeProperties34);
            shape34.Append(shapeProperties34);
            shape34.Append(textBody34);

            shapeTree8.Append(nonVisualGroupShapeProperties8);
            shapeTree8.Append(groupShapeProperties8);
            shapeTree8.Append(shape30);
            shapeTree8.Append(shape31);
            shapeTree8.Append(shape32);
            shapeTree8.Append(shape33);
            shapeTree8.Append(shape34);

            var commonSlideDataExtensionList8 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension8 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId8 = new P14.CreationId() {Val = (UInt32Value) 1181834248U};
            creationId8.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension8.Append(creationId8);

            commonSlideDataExtensionList8.Append(commonSlideDataExtension8);

            commonSlideData8.Append(shapeTree8);
            commonSlideData8.Append(commonSlideDataExtensionList8);

            var colorMapOverride7 = new ColorMapOverride();
            var masterColorMapping7 = new A.MasterColorMapping();

            colorMapOverride7.Append(masterColorMapping7);

            slideLayout6.Append(commonSlideData8);
            slideLayout6.Append(colorMapOverride7);

            slideLayoutPart6.SlideLayout = slideLayout6;
        }

        // Generates content of slideLayoutPart7.
        private static void GenerateSlideLayoutPart7Content(SlideLayoutPart slideLayoutPart7)
        {
            var slideLayout7 = new SlideLayout() {Type = SlideLayoutValues.VerticalTitleAndText, Preserve = true};
            slideLayout7.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout7.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout7.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData9 = new CommonSlideData() {Name = "Vertical Title and Text"};

            var shapeTree9 = new ShapeTree();

            var nonVisualGroupShapeProperties9 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties43 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties9 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties43 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties9.Append(nonVisualDrawingProperties43);
            nonVisualGroupShapeProperties9.Append(nonVisualGroupShapeDrawingProperties9);
            nonVisualGroupShapeProperties9.Append(applicationNonVisualDrawingProperties43);

            var groupShapeProperties9 = new GroupShapeProperties();

            var transformGroup9 = new A.TransformGroup();
            var offset21 = new A.Offset() {X = 0L, Y = 0L};
            var extents21 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset9 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents9 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup9.Append(offset21);
            transformGroup9.Append(extents21);
            transformGroup9.Append(childOffset9);
            transformGroup9.Append(childExtents9);

            groupShapeProperties9.Append(transformGroup9);

            var shape35 = new Shape();

            var nonVisualShapeProperties35 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties44 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 2U,
                Name = "Vertical Title 1"
            };

            var nonVisualShapeDrawingProperties35 = new NonVisualShapeDrawingProperties();
            var shapeLocks35 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties35.Append(shapeLocks35);

            var applicationNonVisualDrawingProperties44 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape35 = new PlaceholderShape() {
                Type = PlaceholderValues.Title,
                Orientation = DirectionValues.Vertical
            };

            applicationNonVisualDrawingProperties44.Append(placeholderShape35);

            nonVisualShapeProperties35.Append(nonVisualDrawingProperties44);
            nonVisualShapeProperties35.Append(nonVisualShapeDrawingProperties35);
            nonVisualShapeProperties35.Append(applicationNonVisualDrawingProperties44);

            var shapeProperties35 = new ShapeProperties();

            var transform2D13 = new A.Transform2D();
            var offset22 = new A.Offset() {X = 6629400L, Y = 274638L};
            var extents22 = new A.Extents() {Cx = 2057400L, Cy = 5851525L};

            transform2D13.Append(offset22);
            transform2D13.Append(extents22);

            shapeProperties35.Append(transform2D13);

            var textBody35 = new TextBody();
            var bodyProperties35 = new A.BodyProperties() {Vertical = A.TextVerticalValues.EastAsianVetical};
            var listStyle35 = new A.ListStyle();

            var paragraph47 = new A.Paragraph();

            var run25 = new A.Run();

            var runProperties39 = new A.RunProperties() {Language = "en-US"};
            runProperties39.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text39 = new A.Text();
            text39.Text = "Click to edit Master title style";

            run25.Append(runProperties39);
            run25.Append(text39);
            var endParagraphRunProperties33 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph47.Append(run25);
            paragraph47.Append(endParagraphRunProperties33);

            textBody35.Append(bodyProperties35);
            textBody35.Append(listStyle35);
            textBody35.Append(paragraph47);

            shape35.Append(nonVisualShapeProperties35);
            shape35.Append(shapeProperties35);
            shape35.Append(textBody35);

            var shape36 = new Shape();

            var nonVisualShapeProperties36 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties45 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Vertical Text Placeholder 2"
            };

            var nonVisualShapeDrawingProperties36 = new NonVisualShapeDrawingProperties();
            var shapeLocks36 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties36.Append(shapeLocks36);

            var applicationNonVisualDrawingProperties45 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape36 = new PlaceholderShape() {
                Type = PlaceholderValues.Body,
                Orientation = DirectionValues.Vertical,
                Index = (UInt32Value) 1U
            };

            applicationNonVisualDrawingProperties45.Append(placeholderShape36);

            nonVisualShapeProperties36.Append(nonVisualDrawingProperties45);
            nonVisualShapeProperties36.Append(nonVisualShapeDrawingProperties36);
            nonVisualShapeProperties36.Append(applicationNonVisualDrawingProperties45);

            var shapeProperties36 = new ShapeProperties();

            var transform2D14 = new A.Transform2D();
            var offset23 = new A.Offset() {X = 457200L, Y = 274638L};
            var extents23 = new A.Extents() {Cx = 6019800L, Cy = 5851525L};

            transform2D14.Append(offset23);
            transform2D14.Append(extents23);

            shapeProperties36.Append(transform2D14);

            var textBody36 = new TextBody();
            var bodyProperties36 = new A.BodyProperties() {Vertical = A.TextVerticalValues.EastAsianVetical};
            var listStyle36 = new A.ListStyle();

            var paragraph48 = new A.Paragraph();
            var paragraphProperties18 = new A.ParagraphProperties() {Level = 0};

            var run26 = new A.Run();

            var runProperties40 = new A.RunProperties() {Language = "en-US"};
            runProperties40.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text40 = new A.Text();
            text40.Text = "Click to edit Master text styles";

            run26.Append(runProperties40);
            run26.Append(text40);

            paragraph48.Append(paragraphProperties18);
            paragraph48.Append(run26);

            var paragraph49 = new A.Paragraph();
            var paragraphProperties19 = new A.ParagraphProperties() {Level = 1};

            var run27 = new A.Run();

            var runProperties41 = new A.RunProperties() {Language = "en-US"};
            runProperties41.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text41 = new A.Text();
            text41.Text = "Second level";

            run27.Append(runProperties41);
            run27.Append(text41);

            paragraph49.Append(paragraphProperties19);
            paragraph49.Append(run27);

            var paragraph50 = new A.Paragraph();
            var paragraphProperties20 = new A.ParagraphProperties() {Level = 2};

            var run28 = new A.Run();

            var runProperties42 = new A.RunProperties() {Language = "en-US"};
            runProperties42.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text42 = new A.Text();
            text42.Text = "Third level";

            run28.Append(runProperties42);
            run28.Append(text42);

            paragraph50.Append(paragraphProperties20);
            paragraph50.Append(run28);

            var paragraph51 = new A.Paragraph();
            var paragraphProperties21 = new A.ParagraphProperties() {Level = 3};

            var run29 = new A.Run();

            var runProperties43 = new A.RunProperties() {Language = "en-US"};
            runProperties43.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text43 = new A.Text();
            text43.Text = "Fourth level";

            run29.Append(runProperties43);
            run29.Append(text43);

            paragraph51.Append(paragraphProperties21);
            paragraph51.Append(run29);

            var paragraph52 = new A.Paragraph();
            var paragraphProperties22 = new A.ParagraphProperties() {Level = 4};

            var run30 = new A.Run();

            var runProperties44 = new A.RunProperties() {Language = "en-US"};
            runProperties44.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text44 = new A.Text();
            text44.Text = "Fifth level";

            run30.Append(runProperties44);
            run30.Append(text44);
            var endParagraphRunProperties34 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph52.Append(paragraphProperties22);
            paragraph52.Append(run30);
            paragraph52.Append(endParagraphRunProperties34);

            textBody36.Append(bodyProperties36);
            textBody36.Append(listStyle36);
            textBody36.Append(paragraph48);
            textBody36.Append(paragraph49);
            textBody36.Append(paragraph50);
            textBody36.Append(paragraph51);
            textBody36.Append(paragraph52);

            shape36.Append(nonVisualShapeProperties36);
            shape36.Append(shapeProperties36);
            shape36.Append(textBody36);

            var shape37 = new Shape();

            var nonVisualShapeProperties37 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties46 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Date Placeholder 3"
            };

            var nonVisualShapeDrawingProperties37 = new NonVisualShapeDrawingProperties();
            var shapeLocks37 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties37.Append(shapeLocks37);

            var applicationNonVisualDrawingProperties46 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape37 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 10U
            };

            applicationNonVisualDrawingProperties46.Append(placeholderShape37);

            nonVisualShapeProperties37.Append(nonVisualDrawingProperties46);
            nonVisualShapeProperties37.Append(nonVisualShapeDrawingProperties37);
            nonVisualShapeProperties37.Append(applicationNonVisualDrawingProperties46);
            var shapeProperties37 = new ShapeProperties();

            var textBody37 = new TextBody();
            var bodyProperties37 = new A.BodyProperties();
            var listStyle37 = new A.ListStyle();

            var paragraph53 = new A.Paragraph();

            var field15 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties45 = new A.RunProperties() {Language = "en-GB"};
            runProperties45.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text45 = new A.Text();
            text45.Text = "04/06/2014";

            field15.Append(runProperties45);
            field15.Append(text45);
            var endParagraphRunProperties35 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph53.Append(field15);
            paragraph53.Append(endParagraphRunProperties35);

            textBody37.Append(bodyProperties37);
            textBody37.Append(listStyle37);
            textBody37.Append(paragraph53);

            shape37.Append(nonVisualShapeProperties37);
            shape37.Append(shapeProperties37);
            shape37.Append(textBody37);

            var shape38 = new Shape();

            var nonVisualShapeProperties38 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties47 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 5U,
                Name = "Footer Placeholder 4"
            };

            var nonVisualShapeDrawingProperties38 = new NonVisualShapeDrawingProperties();
            var shapeLocks38 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties38.Append(shapeLocks38);

            var applicationNonVisualDrawingProperties47 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape38 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 11U
            };

            applicationNonVisualDrawingProperties47.Append(placeholderShape38);

            nonVisualShapeProperties38.Append(nonVisualDrawingProperties47);
            nonVisualShapeProperties38.Append(nonVisualShapeDrawingProperties38);
            nonVisualShapeProperties38.Append(applicationNonVisualDrawingProperties47);
            var shapeProperties38 = new ShapeProperties();

            var textBody38 = new TextBody();
            var bodyProperties38 = new A.BodyProperties();
            var listStyle38 = new A.ListStyle();

            var paragraph54 = new A.Paragraph();
            var endParagraphRunProperties36 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph54.Append(endParagraphRunProperties36);

            textBody38.Append(bodyProperties38);
            textBody38.Append(listStyle38);
            textBody38.Append(paragraph54);

            shape38.Append(nonVisualShapeProperties38);
            shape38.Append(shapeProperties38);
            shape38.Append(textBody38);

            var shape39 = new Shape();

            var nonVisualShapeProperties39 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties48 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 6U,
                Name = "Slide Number Placeholder 5"
            };

            var nonVisualShapeDrawingProperties39 = new NonVisualShapeDrawingProperties();
            var shapeLocks39 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties39.Append(shapeLocks39);

            var applicationNonVisualDrawingProperties48 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape39 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 12U
            };

            applicationNonVisualDrawingProperties48.Append(placeholderShape39);

            nonVisualShapeProperties39.Append(nonVisualDrawingProperties48);
            nonVisualShapeProperties39.Append(nonVisualShapeDrawingProperties39);
            nonVisualShapeProperties39.Append(applicationNonVisualDrawingProperties48);
            var shapeProperties39 = new ShapeProperties();

            var textBody39 = new TextBody();
            var bodyProperties39 = new A.BodyProperties();
            var listStyle39 = new A.ListStyle();

            var paragraph55 = new A.Paragraph();

            var field16 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties46 = new A.RunProperties() {Language = "en-GB"};
            runProperties46.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text46 = new A.Text();
            text46.Text = "‹#›";

            field16.Append(runProperties46);
            field16.Append(text46);
            var endParagraphRunProperties37 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph55.Append(field16);
            paragraph55.Append(endParagraphRunProperties37);

            textBody39.Append(bodyProperties39);
            textBody39.Append(listStyle39);
            textBody39.Append(paragraph55);

            shape39.Append(nonVisualShapeProperties39);
            shape39.Append(shapeProperties39);
            shape39.Append(textBody39);

            shapeTree9.Append(nonVisualGroupShapeProperties9);
            shapeTree9.Append(groupShapeProperties9);
            shapeTree9.Append(shape35);
            shapeTree9.Append(shape36);
            shapeTree9.Append(shape37);
            shapeTree9.Append(shape38);
            shapeTree9.Append(shape39);

            var commonSlideDataExtensionList9 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension9 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId9 = new P14.CreationId() {Val = (UInt32Value) 2049688285U};
            creationId9.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension9.Append(creationId9);

            commonSlideDataExtensionList9.Append(commonSlideDataExtension9);

            commonSlideData9.Append(shapeTree9);
            commonSlideData9.Append(commonSlideDataExtensionList9);

            var colorMapOverride8 = new ColorMapOverride();
            var masterColorMapping8 = new A.MasterColorMapping();

            colorMapOverride8.Append(masterColorMapping8);

            slideLayout7.Append(commonSlideData9);
            slideLayout7.Append(colorMapOverride8);

            slideLayoutPart7.SlideLayout = slideLayout7;
        }

        // Generates content of slideLayoutPart8.
        private static void GenerateSlideLayoutPart8Content(SlideLayoutPart slideLayoutPart8)
        {
            var slideLayout8 = new SlideLayout() {Type = SlideLayoutValues.TwoTextAndTwoObjects, Preserve = true};
            slideLayout8.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout8.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout8.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData10 = new CommonSlideData() {Name = "Comparison"};

            var shapeTree10 = new ShapeTree();

            var nonVisualGroupShapeProperties10 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties49 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties10 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties49 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties10.Append(nonVisualDrawingProperties49);
            nonVisualGroupShapeProperties10.Append(nonVisualGroupShapeDrawingProperties10);
            nonVisualGroupShapeProperties10.Append(applicationNonVisualDrawingProperties49);

            var groupShapeProperties10 = new GroupShapeProperties();

            var transformGroup10 = new A.TransformGroup();
            var offset24 = new A.Offset() {X = 0L, Y = 0L};
            var extents24 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset10 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents10 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup10.Append(offset24);
            transformGroup10.Append(extents24);
            transformGroup10.Append(childOffset10);
            transformGroup10.Append(childExtents10);

            groupShapeProperties10.Append(transformGroup10);

            var shape40 = new Shape();

            var nonVisualShapeProperties40 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties50 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 2U,
                Name = "Title 1"
            };

            var nonVisualShapeDrawingProperties40 = new NonVisualShapeDrawingProperties();
            var shapeLocks40 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties40.Append(shapeLocks40);

            var applicationNonVisualDrawingProperties50 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape40 = new PlaceholderShape() {Type = PlaceholderValues.Title};

            applicationNonVisualDrawingProperties50.Append(placeholderShape40);

            nonVisualShapeProperties40.Append(nonVisualDrawingProperties50);
            nonVisualShapeProperties40.Append(nonVisualShapeDrawingProperties40);
            nonVisualShapeProperties40.Append(applicationNonVisualDrawingProperties50);
            var shapeProperties40 = new ShapeProperties();

            var textBody40 = new TextBody();
            var bodyProperties40 = new A.BodyProperties();

            var listStyle40 = new A.ListStyle();

            var level1ParagraphProperties14 = new A.Level1ParagraphProperties();
            var defaultRunProperties72 = new A.DefaultRunProperties();

            level1ParagraphProperties14.Append(defaultRunProperties72);

            listStyle40.Append(level1ParagraphProperties14);

            var paragraph56 = new A.Paragraph();

            var run31 = new A.Run();

            var runProperties47 = new A.RunProperties() {Language = "en-US"};
            runProperties47.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text47 = new A.Text();
            text47.Text = "Click to edit Master title style";

            run31.Append(runProperties47);
            run31.Append(text47);
            var endParagraphRunProperties38 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph56.Append(run31);
            paragraph56.Append(endParagraphRunProperties38);

            textBody40.Append(bodyProperties40);
            textBody40.Append(listStyle40);
            textBody40.Append(paragraph56);

            shape40.Append(nonVisualShapeProperties40);
            shape40.Append(shapeProperties40);
            shape40.Append(textBody40);

            var shape41 = new Shape();

            var nonVisualShapeProperties41 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties51 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Text Placeholder 2"
            };

            var nonVisualShapeDrawingProperties41 = new NonVisualShapeDrawingProperties();
            var shapeLocks41 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties41.Append(shapeLocks41);

            var applicationNonVisualDrawingProperties51 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape41 = new PlaceholderShape() {Type = PlaceholderValues.Body, Index = (UInt32Value) 1U};

            applicationNonVisualDrawingProperties51.Append(placeholderShape41);

            nonVisualShapeProperties41.Append(nonVisualDrawingProperties51);
            nonVisualShapeProperties41.Append(nonVisualShapeDrawingProperties41);
            nonVisualShapeProperties41.Append(applicationNonVisualDrawingProperties51);

            var shapeProperties41 = new ShapeProperties();

            var transform2D15 = new A.Transform2D();
            var offset25 = new A.Offset() {X = 457200L, Y = 1535113L};
            var extents25 = new A.Extents() {Cx = 4040188L, Cy = 639762L};

            transform2D15.Append(offset25);
            transform2D15.Append(extents25);

            shapeProperties41.Append(transform2D15);

            var textBody41 = new TextBody();
            var bodyProperties41 = new A.BodyProperties() {Anchor = A.TextAnchoringTypeValues.Bottom};

            var listStyle41 = new A.ListStyle();

            var level1ParagraphProperties15 = new A.Level1ParagraphProperties() {LeftMargin = 0, Indent = 0};
            var noBullet29 = new A.NoBullet();
            var defaultRunProperties73 = new A.DefaultRunProperties() {FontSize = 2400, Bold = true};

            level1ParagraphProperties15.Append(noBullet29);
            level1ParagraphProperties15.Append(defaultRunProperties73);

            var level2ParagraphProperties8 = new A.Level2ParagraphProperties() {LeftMargin = 457200, Indent = 0};
            var noBullet30 = new A.NoBullet();
            var defaultRunProperties74 = new A.DefaultRunProperties() {FontSize = 2000, Bold = true};

            level2ParagraphProperties8.Append(noBullet30);
            level2ParagraphProperties8.Append(defaultRunProperties74);

            var level3ParagraphProperties8 = new A.Level3ParagraphProperties() {LeftMargin = 914400, Indent = 0};
            var noBullet31 = new A.NoBullet();
            var defaultRunProperties75 = new A.DefaultRunProperties() {FontSize = 1800, Bold = true};

            level3ParagraphProperties8.Append(noBullet31);
            level3ParagraphProperties8.Append(defaultRunProperties75);

            var level4ParagraphProperties8 = new A.Level4ParagraphProperties() {LeftMargin = 1371600, Indent = 0};
            var noBullet32 = new A.NoBullet();
            var defaultRunProperties76 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level4ParagraphProperties8.Append(noBullet32);
            level4ParagraphProperties8.Append(defaultRunProperties76);

            var level5ParagraphProperties8 = new A.Level5ParagraphProperties() {LeftMargin = 1828800, Indent = 0};
            var noBullet33 = new A.NoBullet();
            var defaultRunProperties77 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level5ParagraphProperties8.Append(noBullet33);
            level5ParagraphProperties8.Append(defaultRunProperties77);

            var level6ParagraphProperties8 = new A.Level6ParagraphProperties() {LeftMargin = 2286000, Indent = 0};
            var noBullet34 = new A.NoBullet();
            var defaultRunProperties78 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level6ParagraphProperties8.Append(noBullet34);
            level6ParagraphProperties8.Append(defaultRunProperties78);

            var level7ParagraphProperties8 = new A.Level7ParagraphProperties() {LeftMargin = 2743200, Indent = 0};
            var noBullet35 = new A.NoBullet();
            var defaultRunProperties79 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level7ParagraphProperties8.Append(noBullet35);
            level7ParagraphProperties8.Append(defaultRunProperties79);

            var level8ParagraphProperties8 = new A.Level8ParagraphProperties() {LeftMargin = 3200400, Indent = 0};
            var noBullet36 = new A.NoBullet();
            var defaultRunProperties80 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level8ParagraphProperties8.Append(noBullet36);
            level8ParagraphProperties8.Append(defaultRunProperties80);

            var level9ParagraphProperties8 = new A.Level9ParagraphProperties() {LeftMargin = 3657600, Indent = 0};
            var noBullet37 = new A.NoBullet();
            var defaultRunProperties81 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level9ParagraphProperties8.Append(noBullet37);
            level9ParagraphProperties8.Append(defaultRunProperties81);

            listStyle41.Append(level1ParagraphProperties15);
            listStyle41.Append(level2ParagraphProperties8);
            listStyle41.Append(level3ParagraphProperties8);
            listStyle41.Append(level4ParagraphProperties8);
            listStyle41.Append(level5ParagraphProperties8);
            listStyle41.Append(level6ParagraphProperties8);
            listStyle41.Append(level7ParagraphProperties8);
            listStyle41.Append(level8ParagraphProperties8);
            listStyle41.Append(level9ParagraphProperties8);

            var paragraph57 = new A.Paragraph();
            var paragraphProperties23 = new A.ParagraphProperties() {Level = 0};

            var run32 = new A.Run();

            var runProperties48 = new A.RunProperties() {Language = "en-US"};
            runProperties48.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text48 = new A.Text();
            text48.Text = "Click to edit Master text styles";

            run32.Append(runProperties48);
            run32.Append(text48);

            paragraph57.Append(paragraphProperties23);
            paragraph57.Append(run32);

            textBody41.Append(bodyProperties41);
            textBody41.Append(listStyle41);
            textBody41.Append(paragraph57);

            shape41.Append(nonVisualShapeProperties41);
            shape41.Append(shapeProperties41);
            shape41.Append(textBody41);

            var shape42 = new Shape();

            var nonVisualShapeProperties42 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties52 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Content Placeholder 3"
            };

            var nonVisualShapeDrawingProperties42 = new NonVisualShapeDrawingProperties();
            var shapeLocks42 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties42.Append(shapeLocks42);

            var applicationNonVisualDrawingProperties52 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape42 = new PlaceholderShape() {
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 2U
            };

            applicationNonVisualDrawingProperties52.Append(placeholderShape42);

            nonVisualShapeProperties42.Append(nonVisualDrawingProperties52);
            nonVisualShapeProperties42.Append(nonVisualShapeDrawingProperties42);
            nonVisualShapeProperties42.Append(applicationNonVisualDrawingProperties52);

            var shapeProperties42 = new ShapeProperties();

            var transform2D16 = new A.Transform2D();
            var offset26 = new A.Offset() {X = 457200L, Y = 2174875L};
            var extents26 = new A.Extents() {Cx = 4040188L, Cy = 3951288L};

            transform2D16.Append(offset26);
            transform2D16.Append(extents26);

            shapeProperties42.Append(transform2D16);

            var textBody42 = new TextBody();
            var bodyProperties42 = new A.BodyProperties();

            var listStyle42 = new A.ListStyle();

            var level1ParagraphProperties16 = new A.Level1ParagraphProperties();
            var defaultRunProperties82 = new A.DefaultRunProperties() {FontSize = 2400};

            level1ParagraphProperties16.Append(defaultRunProperties82);

            var level2ParagraphProperties9 = new A.Level2ParagraphProperties();
            var defaultRunProperties83 = new A.DefaultRunProperties() {FontSize = 2000};

            level2ParagraphProperties9.Append(defaultRunProperties83);

            var level3ParagraphProperties9 = new A.Level3ParagraphProperties();
            var defaultRunProperties84 = new A.DefaultRunProperties() {FontSize = 1800};

            level3ParagraphProperties9.Append(defaultRunProperties84);

            var level4ParagraphProperties9 = new A.Level4ParagraphProperties();
            var defaultRunProperties85 = new A.DefaultRunProperties() {FontSize = 1600};

            level4ParagraphProperties9.Append(defaultRunProperties85);

            var level5ParagraphProperties9 = new A.Level5ParagraphProperties();
            var defaultRunProperties86 = new A.DefaultRunProperties() {FontSize = 1600};

            level5ParagraphProperties9.Append(defaultRunProperties86);

            var level6ParagraphProperties9 = new A.Level6ParagraphProperties();
            var defaultRunProperties87 = new A.DefaultRunProperties() {FontSize = 1600};

            level6ParagraphProperties9.Append(defaultRunProperties87);

            var level7ParagraphProperties9 = new A.Level7ParagraphProperties();
            var defaultRunProperties88 = new A.DefaultRunProperties() {FontSize = 1600};

            level7ParagraphProperties9.Append(defaultRunProperties88);

            var level8ParagraphProperties9 = new A.Level8ParagraphProperties();
            var defaultRunProperties89 = new A.DefaultRunProperties() {FontSize = 1600};

            level8ParagraphProperties9.Append(defaultRunProperties89);

            var level9ParagraphProperties9 = new A.Level9ParagraphProperties();
            var defaultRunProperties90 = new A.DefaultRunProperties() {FontSize = 1600};

            level9ParagraphProperties9.Append(defaultRunProperties90);

            listStyle42.Append(level1ParagraphProperties16);
            listStyle42.Append(level2ParagraphProperties9);
            listStyle42.Append(level3ParagraphProperties9);
            listStyle42.Append(level4ParagraphProperties9);
            listStyle42.Append(level5ParagraphProperties9);
            listStyle42.Append(level6ParagraphProperties9);
            listStyle42.Append(level7ParagraphProperties9);
            listStyle42.Append(level8ParagraphProperties9);
            listStyle42.Append(level9ParagraphProperties9);

            var paragraph58 = new A.Paragraph();
            var paragraphProperties24 = new A.ParagraphProperties() {Level = 0};

            var run33 = new A.Run();

            var runProperties49 = new A.RunProperties() {Language = "en-US"};
            runProperties49.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text49 = new A.Text();
            text49.Text = "Click to edit Master text styles";

            run33.Append(runProperties49);
            run33.Append(text49);

            paragraph58.Append(paragraphProperties24);
            paragraph58.Append(run33);

            var paragraph59 = new A.Paragraph();
            var paragraphProperties25 = new A.ParagraphProperties() {Level = 1};

            var run34 = new A.Run();

            var runProperties50 = new A.RunProperties() {Language = "en-US"};
            runProperties50.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text50 = new A.Text();
            text50.Text = "Second level";

            run34.Append(runProperties50);
            run34.Append(text50);

            paragraph59.Append(paragraphProperties25);
            paragraph59.Append(run34);

            var paragraph60 = new A.Paragraph();
            var paragraphProperties26 = new A.ParagraphProperties() {Level = 2};

            var run35 = new A.Run();

            var runProperties51 = new A.RunProperties() {Language = "en-US"};
            runProperties51.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text51 = new A.Text();
            text51.Text = "Third level";

            run35.Append(runProperties51);
            run35.Append(text51);

            paragraph60.Append(paragraphProperties26);
            paragraph60.Append(run35);

            var paragraph61 = new A.Paragraph();
            var paragraphProperties27 = new A.ParagraphProperties() {Level = 3};

            var run36 = new A.Run();

            var runProperties52 = new A.RunProperties() {Language = "en-US"};
            runProperties52.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text52 = new A.Text();
            text52.Text = "Fourth level";

            run36.Append(runProperties52);
            run36.Append(text52);

            paragraph61.Append(paragraphProperties27);
            paragraph61.Append(run36);

            var paragraph62 = new A.Paragraph();
            var paragraphProperties28 = new A.ParagraphProperties() {Level = 4};

            var run37 = new A.Run();

            var runProperties53 = new A.RunProperties() {Language = "en-US"};
            runProperties53.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text53 = new A.Text();
            text53.Text = "Fifth level";

            run37.Append(runProperties53);
            run37.Append(text53);
            var endParagraphRunProperties39 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph62.Append(paragraphProperties28);
            paragraph62.Append(run37);
            paragraph62.Append(endParagraphRunProperties39);

            textBody42.Append(bodyProperties42);
            textBody42.Append(listStyle42);
            textBody42.Append(paragraph58);
            textBody42.Append(paragraph59);
            textBody42.Append(paragraph60);
            textBody42.Append(paragraph61);
            textBody42.Append(paragraph62);

            shape42.Append(nonVisualShapeProperties42);
            shape42.Append(shapeProperties42);
            shape42.Append(textBody42);

            var shape43 = new Shape();

            var nonVisualShapeProperties43 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties53 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 5U,
                Name = "Text Placeholder 4"
            };

            var nonVisualShapeDrawingProperties43 = new NonVisualShapeDrawingProperties();
            var shapeLocks43 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties43.Append(shapeLocks43);

            var applicationNonVisualDrawingProperties53 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape43 = new PlaceholderShape() {
                Type = PlaceholderValues.Body,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 3U
            };

            applicationNonVisualDrawingProperties53.Append(placeholderShape43);

            nonVisualShapeProperties43.Append(nonVisualDrawingProperties53);
            nonVisualShapeProperties43.Append(nonVisualShapeDrawingProperties43);
            nonVisualShapeProperties43.Append(applicationNonVisualDrawingProperties53);

            var shapeProperties43 = new ShapeProperties();

            var transform2D17 = new A.Transform2D();
            var offset27 = new A.Offset() {X = 4645025L, Y = 1535113L};
            var extents27 = new A.Extents() {Cx = 4041775L, Cy = 639762L};

            transform2D17.Append(offset27);
            transform2D17.Append(extents27);

            shapeProperties43.Append(transform2D17);

            var textBody43 = new TextBody();
            var bodyProperties43 = new A.BodyProperties() {Anchor = A.TextAnchoringTypeValues.Bottom};

            var listStyle43 = new A.ListStyle();

            var level1ParagraphProperties17 = new A.Level1ParagraphProperties() {LeftMargin = 0, Indent = 0};
            var noBullet38 = new A.NoBullet();
            var defaultRunProperties91 = new A.DefaultRunProperties() {FontSize = 2400, Bold = true};

            level1ParagraphProperties17.Append(noBullet38);
            level1ParagraphProperties17.Append(defaultRunProperties91);

            var level2ParagraphProperties10 = new A.Level2ParagraphProperties() {LeftMargin = 457200, Indent = 0};
            var noBullet39 = new A.NoBullet();
            var defaultRunProperties92 = new A.DefaultRunProperties() {FontSize = 2000, Bold = true};

            level2ParagraphProperties10.Append(noBullet39);
            level2ParagraphProperties10.Append(defaultRunProperties92);

            var level3ParagraphProperties10 = new A.Level3ParagraphProperties() {LeftMargin = 914400, Indent = 0};
            var noBullet40 = new A.NoBullet();
            var defaultRunProperties93 = new A.DefaultRunProperties() {FontSize = 1800, Bold = true};

            level3ParagraphProperties10.Append(noBullet40);
            level3ParagraphProperties10.Append(defaultRunProperties93);

            var level4ParagraphProperties10 = new A.Level4ParagraphProperties() {LeftMargin = 1371600, Indent = 0};
            var noBullet41 = new A.NoBullet();
            var defaultRunProperties94 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level4ParagraphProperties10.Append(noBullet41);
            level4ParagraphProperties10.Append(defaultRunProperties94);

            var level5ParagraphProperties10 = new A.Level5ParagraphProperties() {LeftMargin = 1828800, Indent = 0};
            var noBullet42 = new A.NoBullet();
            var defaultRunProperties95 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level5ParagraphProperties10.Append(noBullet42);
            level5ParagraphProperties10.Append(defaultRunProperties95);

            var level6ParagraphProperties10 = new A.Level6ParagraphProperties() {LeftMargin = 2286000, Indent = 0};
            var noBullet43 = new A.NoBullet();
            var defaultRunProperties96 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level6ParagraphProperties10.Append(noBullet43);
            level6ParagraphProperties10.Append(defaultRunProperties96);

            var level7ParagraphProperties10 = new A.Level7ParagraphProperties() {LeftMargin = 2743200, Indent = 0};
            var noBullet44 = new A.NoBullet();
            var defaultRunProperties97 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level7ParagraphProperties10.Append(noBullet44);
            level7ParagraphProperties10.Append(defaultRunProperties97);

            var level8ParagraphProperties10 = new A.Level8ParagraphProperties() {LeftMargin = 3200400, Indent = 0};
            var noBullet45 = new A.NoBullet();
            var defaultRunProperties98 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level8ParagraphProperties10.Append(noBullet45);
            level8ParagraphProperties10.Append(defaultRunProperties98);

            var level9ParagraphProperties10 = new A.Level9ParagraphProperties() {LeftMargin = 3657600, Indent = 0};
            var noBullet46 = new A.NoBullet();
            var defaultRunProperties99 = new A.DefaultRunProperties() {FontSize = 1600, Bold = true};

            level9ParagraphProperties10.Append(noBullet46);
            level9ParagraphProperties10.Append(defaultRunProperties99);

            listStyle43.Append(level1ParagraphProperties17);
            listStyle43.Append(level2ParagraphProperties10);
            listStyle43.Append(level3ParagraphProperties10);
            listStyle43.Append(level4ParagraphProperties10);
            listStyle43.Append(level5ParagraphProperties10);
            listStyle43.Append(level6ParagraphProperties10);
            listStyle43.Append(level7ParagraphProperties10);
            listStyle43.Append(level8ParagraphProperties10);
            listStyle43.Append(level9ParagraphProperties10);

            var paragraph63 = new A.Paragraph();
            var paragraphProperties29 = new A.ParagraphProperties() {Level = 0};

            var run38 = new A.Run();

            var runProperties54 = new A.RunProperties() {Language = "en-US"};
            runProperties54.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text54 = new A.Text();
            text54.Text = "Click to edit Master text styles";

            run38.Append(runProperties54);
            run38.Append(text54);

            paragraph63.Append(paragraphProperties29);
            paragraph63.Append(run38);

            textBody43.Append(bodyProperties43);
            textBody43.Append(listStyle43);
            textBody43.Append(paragraph63);

            shape43.Append(nonVisualShapeProperties43);
            shape43.Append(shapeProperties43);
            shape43.Append(textBody43);

            var shape44 = new Shape();

            var nonVisualShapeProperties44 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties54 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 6U,
                Name = "Content Placeholder 5"
            };

            var nonVisualShapeDrawingProperties44 = new NonVisualShapeDrawingProperties();
            var shapeLocks44 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties44.Append(shapeLocks44);

            var applicationNonVisualDrawingProperties54 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape44 = new PlaceholderShape() {
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 4U
            };

            applicationNonVisualDrawingProperties54.Append(placeholderShape44);

            nonVisualShapeProperties44.Append(nonVisualDrawingProperties54);
            nonVisualShapeProperties44.Append(nonVisualShapeDrawingProperties44);
            nonVisualShapeProperties44.Append(applicationNonVisualDrawingProperties54);

            var shapeProperties44 = new ShapeProperties();

            var transform2D18 = new A.Transform2D();
            var offset28 = new A.Offset() {X = 4645025L, Y = 2174875L};
            var extents28 = new A.Extents() {Cx = 4041775L, Cy = 3951288L};

            transform2D18.Append(offset28);
            transform2D18.Append(extents28);

            shapeProperties44.Append(transform2D18);

            var textBody44 = new TextBody();
            var bodyProperties44 = new A.BodyProperties();

            var listStyle44 = new A.ListStyle();

            var level1ParagraphProperties18 = new A.Level1ParagraphProperties();
            var defaultRunProperties100 = new A.DefaultRunProperties() {FontSize = 2400};

            level1ParagraphProperties18.Append(defaultRunProperties100);

            var level2ParagraphProperties11 = new A.Level2ParagraphProperties();
            var defaultRunProperties101 = new A.DefaultRunProperties() {FontSize = 2000};

            level2ParagraphProperties11.Append(defaultRunProperties101);

            var level3ParagraphProperties11 = new A.Level3ParagraphProperties();
            var defaultRunProperties102 = new A.DefaultRunProperties() {FontSize = 1800};

            level3ParagraphProperties11.Append(defaultRunProperties102);

            var level4ParagraphProperties11 = new A.Level4ParagraphProperties();
            var defaultRunProperties103 = new A.DefaultRunProperties() {FontSize = 1600};

            level4ParagraphProperties11.Append(defaultRunProperties103);

            var level5ParagraphProperties11 = new A.Level5ParagraphProperties();
            var defaultRunProperties104 = new A.DefaultRunProperties() {FontSize = 1600};

            level5ParagraphProperties11.Append(defaultRunProperties104);

            var level6ParagraphProperties11 = new A.Level6ParagraphProperties();
            var defaultRunProperties105 = new A.DefaultRunProperties() {FontSize = 1600};

            level6ParagraphProperties11.Append(defaultRunProperties105);

            var level7ParagraphProperties11 = new A.Level7ParagraphProperties();
            var defaultRunProperties106 = new A.DefaultRunProperties() {FontSize = 1600};

            level7ParagraphProperties11.Append(defaultRunProperties106);

            var level8ParagraphProperties11 = new A.Level8ParagraphProperties();
            var defaultRunProperties107 = new A.DefaultRunProperties() {FontSize = 1600};

            level8ParagraphProperties11.Append(defaultRunProperties107);

            var level9ParagraphProperties11 = new A.Level9ParagraphProperties();
            var defaultRunProperties108 = new A.DefaultRunProperties() {FontSize = 1600};

            level9ParagraphProperties11.Append(defaultRunProperties108);

            listStyle44.Append(level1ParagraphProperties18);
            listStyle44.Append(level2ParagraphProperties11);
            listStyle44.Append(level3ParagraphProperties11);
            listStyle44.Append(level4ParagraphProperties11);
            listStyle44.Append(level5ParagraphProperties11);
            listStyle44.Append(level6ParagraphProperties11);
            listStyle44.Append(level7ParagraphProperties11);
            listStyle44.Append(level8ParagraphProperties11);
            listStyle44.Append(level9ParagraphProperties11);

            var paragraph64 = new A.Paragraph();
            var paragraphProperties30 = new A.ParagraphProperties() {Level = 0};

            var run39 = new A.Run();

            var runProperties55 = new A.RunProperties() {Language = "en-US"};
            runProperties55.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text55 = new A.Text();
            text55.Text = "Click to edit Master text styles";

            run39.Append(runProperties55);
            run39.Append(text55);

            paragraph64.Append(paragraphProperties30);
            paragraph64.Append(run39);

            var paragraph65 = new A.Paragraph();
            var paragraphProperties31 = new A.ParagraphProperties() {Level = 1};

            var run40 = new A.Run();

            var runProperties56 = new A.RunProperties() {Language = "en-US"};
            runProperties56.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text56 = new A.Text();
            text56.Text = "Second level";

            run40.Append(runProperties56);
            run40.Append(text56);

            paragraph65.Append(paragraphProperties31);
            paragraph65.Append(run40);

            var paragraph66 = new A.Paragraph();
            var paragraphProperties32 = new A.ParagraphProperties() {Level = 2};

            var run41 = new A.Run();

            var runProperties57 = new A.RunProperties() {Language = "en-US"};
            runProperties57.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text57 = new A.Text();
            text57.Text = "Third level";

            run41.Append(runProperties57);
            run41.Append(text57);

            paragraph66.Append(paragraphProperties32);
            paragraph66.Append(run41);

            var paragraph67 = new A.Paragraph();
            var paragraphProperties33 = new A.ParagraphProperties() {Level = 3};

            var run42 = new A.Run();

            var runProperties58 = new A.RunProperties() {Language = "en-US"};
            runProperties58.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text58 = new A.Text();
            text58.Text = "Fourth level";

            run42.Append(runProperties58);
            run42.Append(text58);

            paragraph67.Append(paragraphProperties33);
            paragraph67.Append(run42);

            var paragraph68 = new A.Paragraph();
            var paragraphProperties34 = new A.ParagraphProperties() {Level = 4};

            var run43 = new A.Run();

            var runProperties59 = new A.RunProperties() {Language = "en-US"};
            runProperties59.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text59 = new A.Text();
            text59.Text = "Fifth level";

            run43.Append(runProperties59);
            run43.Append(text59);
            var endParagraphRunProperties40 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph68.Append(paragraphProperties34);
            paragraph68.Append(run43);
            paragraph68.Append(endParagraphRunProperties40);

            textBody44.Append(bodyProperties44);
            textBody44.Append(listStyle44);
            textBody44.Append(paragraph64);
            textBody44.Append(paragraph65);
            textBody44.Append(paragraph66);
            textBody44.Append(paragraph67);
            textBody44.Append(paragraph68);

            shape44.Append(nonVisualShapeProperties44);
            shape44.Append(shapeProperties44);
            shape44.Append(textBody44);

            var shape45 = new Shape();

            var nonVisualShapeProperties45 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties55 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 7U,
                Name = "Date Placeholder 6"
            };

            var nonVisualShapeDrawingProperties45 = new NonVisualShapeDrawingProperties();
            var shapeLocks45 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties45.Append(shapeLocks45);

            var applicationNonVisualDrawingProperties55 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape45 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 10U
            };

            applicationNonVisualDrawingProperties55.Append(placeholderShape45);

            nonVisualShapeProperties45.Append(nonVisualDrawingProperties55);
            nonVisualShapeProperties45.Append(nonVisualShapeDrawingProperties45);
            nonVisualShapeProperties45.Append(applicationNonVisualDrawingProperties55);
            var shapeProperties45 = new ShapeProperties();

            var textBody45 = new TextBody();
            var bodyProperties45 = new A.BodyProperties();
            var listStyle45 = new A.ListStyle();

            var paragraph69 = new A.Paragraph();

            var field17 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties60 = new A.RunProperties() {Language = "en-GB"};
            runProperties60.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text60 = new A.Text();
            text60.Text = "04/06/2014";

            field17.Append(runProperties60);
            field17.Append(text60);
            var endParagraphRunProperties41 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph69.Append(field17);
            paragraph69.Append(endParagraphRunProperties41);

            textBody45.Append(bodyProperties45);
            textBody45.Append(listStyle45);
            textBody45.Append(paragraph69);

            shape45.Append(nonVisualShapeProperties45);
            shape45.Append(shapeProperties45);
            shape45.Append(textBody45);

            var shape46 = new Shape();

            var nonVisualShapeProperties46 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties56 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 8U,
                Name = "Footer Placeholder 7"
            };

            var nonVisualShapeDrawingProperties46 = new NonVisualShapeDrawingProperties();
            var shapeLocks46 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties46.Append(shapeLocks46);

            var applicationNonVisualDrawingProperties56 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape46 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 11U
            };

            applicationNonVisualDrawingProperties56.Append(placeholderShape46);

            nonVisualShapeProperties46.Append(nonVisualDrawingProperties56);
            nonVisualShapeProperties46.Append(nonVisualShapeDrawingProperties46);
            nonVisualShapeProperties46.Append(applicationNonVisualDrawingProperties56);
            var shapeProperties46 = new ShapeProperties();

            var textBody46 = new TextBody();
            var bodyProperties46 = new A.BodyProperties();
            var listStyle46 = new A.ListStyle();

            var paragraph70 = new A.Paragraph();
            var endParagraphRunProperties42 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph70.Append(endParagraphRunProperties42);

            textBody46.Append(bodyProperties46);
            textBody46.Append(listStyle46);
            textBody46.Append(paragraph70);

            shape46.Append(nonVisualShapeProperties46);
            shape46.Append(shapeProperties46);
            shape46.Append(textBody46);

            var shape47 = new Shape();

            var nonVisualShapeProperties47 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties57 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 9U,
                Name = "Slide Number Placeholder 8"
            };

            var nonVisualShapeDrawingProperties47 = new NonVisualShapeDrawingProperties();
            var shapeLocks47 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties47.Append(shapeLocks47);

            var applicationNonVisualDrawingProperties57 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape47 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 12U
            };

            applicationNonVisualDrawingProperties57.Append(placeholderShape47);

            nonVisualShapeProperties47.Append(nonVisualDrawingProperties57);
            nonVisualShapeProperties47.Append(nonVisualShapeDrawingProperties47);
            nonVisualShapeProperties47.Append(applicationNonVisualDrawingProperties57);
            var shapeProperties47 = new ShapeProperties();

            var textBody47 = new TextBody();
            var bodyProperties47 = new A.BodyProperties();
            var listStyle47 = new A.ListStyle();

            var paragraph71 = new A.Paragraph();

            var field18 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties61 = new A.RunProperties() {Language = "en-GB"};
            runProperties61.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text61 = new A.Text();
            text61.Text = "‹#›";

            field18.Append(runProperties61);
            field18.Append(text61);
            var endParagraphRunProperties43 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph71.Append(field18);
            paragraph71.Append(endParagraphRunProperties43);

            textBody47.Append(bodyProperties47);
            textBody47.Append(listStyle47);
            textBody47.Append(paragraph71);

            shape47.Append(nonVisualShapeProperties47);
            shape47.Append(shapeProperties47);
            shape47.Append(textBody47);

            shapeTree10.Append(nonVisualGroupShapeProperties10);
            shapeTree10.Append(groupShapeProperties10);
            shapeTree10.Append(shape40);
            shapeTree10.Append(shape41);
            shapeTree10.Append(shape42);
            shapeTree10.Append(shape43);
            shapeTree10.Append(shape44);
            shapeTree10.Append(shape45);
            shapeTree10.Append(shape46);
            shapeTree10.Append(shape47);

            var commonSlideDataExtensionList10 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension10 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId10 = new P14.CreationId() {Val = (UInt32Value) 3596975214U};
            creationId10.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension10.Append(creationId10);

            commonSlideDataExtensionList10.Append(commonSlideDataExtension10);

            commonSlideData10.Append(shapeTree10);
            commonSlideData10.Append(commonSlideDataExtensionList10);

            var colorMapOverride9 = new ColorMapOverride();
            var masterColorMapping9 = new A.MasterColorMapping();

            colorMapOverride9.Append(masterColorMapping9);

            slideLayout8.Append(commonSlideData10);
            slideLayout8.Append(colorMapOverride9);

            slideLayoutPart8.SlideLayout = slideLayout8;
        }

        // Generates content of slideLayoutPart9.
        private static void GenerateSlideLayoutPart9Content(SlideLayoutPart slideLayoutPart9)
        {
            var slideLayout9 = new SlideLayout() {Type = SlideLayoutValues.VerticalText, Preserve = true};
            slideLayout9.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout9.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout9.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData11 = new CommonSlideData() {Name = "Title and Vertical Text"};

            var shapeTree11 = new ShapeTree();

            var nonVisualGroupShapeProperties11 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties58 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties11 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties58 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties11.Append(nonVisualDrawingProperties58);
            nonVisualGroupShapeProperties11.Append(nonVisualGroupShapeDrawingProperties11);
            nonVisualGroupShapeProperties11.Append(applicationNonVisualDrawingProperties58);

            var groupShapeProperties11 = new GroupShapeProperties();

            var transformGroup11 = new A.TransformGroup();
            var offset29 = new A.Offset() {X = 0L, Y = 0L};
            var extents29 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset11 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents11 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup11.Append(offset29);
            transformGroup11.Append(extents29);
            transformGroup11.Append(childOffset11);
            transformGroup11.Append(childExtents11);

            groupShapeProperties11.Append(transformGroup11);

            var shape48 = new Shape();

            var nonVisualShapeProperties48 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties59 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 2U,
                Name = "Title 1"
            };

            var nonVisualShapeDrawingProperties48 = new NonVisualShapeDrawingProperties();
            var shapeLocks48 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties48.Append(shapeLocks48);

            var applicationNonVisualDrawingProperties59 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape48 = new PlaceholderShape() {Type = PlaceholderValues.Title};

            applicationNonVisualDrawingProperties59.Append(placeholderShape48);

            nonVisualShapeProperties48.Append(nonVisualDrawingProperties59);
            nonVisualShapeProperties48.Append(nonVisualShapeDrawingProperties48);
            nonVisualShapeProperties48.Append(applicationNonVisualDrawingProperties59);
            var shapeProperties48 = new ShapeProperties();

            var textBody48 = new TextBody();
            var bodyProperties48 = new A.BodyProperties();
            var listStyle48 = new A.ListStyle();

            var paragraph72 = new A.Paragraph();

            var run44 = new A.Run();

            var runProperties62 = new A.RunProperties() {Language = "en-US"};
            runProperties62.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text62 = new A.Text();
            text62.Text = "Click to edit Master title style";

            run44.Append(runProperties62);
            run44.Append(text62);
            var endParagraphRunProperties44 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph72.Append(run44);
            paragraph72.Append(endParagraphRunProperties44);

            textBody48.Append(bodyProperties48);
            textBody48.Append(listStyle48);
            textBody48.Append(paragraph72);

            shape48.Append(nonVisualShapeProperties48);
            shape48.Append(shapeProperties48);
            shape48.Append(textBody48);

            var shape49 = new Shape();

            var nonVisualShapeProperties49 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties60 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Vertical Text Placeholder 2"
            };

            var nonVisualShapeDrawingProperties49 = new NonVisualShapeDrawingProperties();
            var shapeLocks49 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties49.Append(shapeLocks49);

            var applicationNonVisualDrawingProperties60 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape49 = new PlaceholderShape() {
                Type = PlaceholderValues.Body,
                Orientation = DirectionValues.Vertical,
                Index = (UInt32Value) 1U
            };

            applicationNonVisualDrawingProperties60.Append(placeholderShape49);

            nonVisualShapeProperties49.Append(nonVisualDrawingProperties60);
            nonVisualShapeProperties49.Append(nonVisualShapeDrawingProperties49);
            nonVisualShapeProperties49.Append(applicationNonVisualDrawingProperties60);
            var shapeProperties49 = new ShapeProperties();

            var textBody49 = new TextBody();
            var bodyProperties49 = new A.BodyProperties() {Vertical = A.TextVerticalValues.EastAsianVetical};
            var listStyle49 = new A.ListStyle();

            var paragraph73 = new A.Paragraph();
            var paragraphProperties35 = new A.ParagraphProperties() {Level = 0};

            var run45 = new A.Run();

            var runProperties63 = new A.RunProperties() {Language = "en-US"};
            runProperties63.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text63 = new A.Text();
            text63.Text = "Click to edit Master text styles";

            run45.Append(runProperties63);
            run45.Append(text63);

            paragraph73.Append(paragraphProperties35);
            paragraph73.Append(run45);

            var paragraph74 = new A.Paragraph();
            var paragraphProperties36 = new A.ParagraphProperties() {Level = 1};

            var run46 = new A.Run();

            var runProperties64 = new A.RunProperties() {Language = "en-US"};
            runProperties64.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text64 = new A.Text();
            text64.Text = "Second level";

            run46.Append(runProperties64);
            run46.Append(text64);

            paragraph74.Append(paragraphProperties36);
            paragraph74.Append(run46);

            var paragraph75 = new A.Paragraph();
            var paragraphProperties37 = new A.ParagraphProperties() {Level = 2};

            var run47 = new A.Run();

            var runProperties65 = new A.RunProperties() {Language = "en-US"};
            runProperties65.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text65 = new A.Text();
            text65.Text = "Third level";

            run47.Append(runProperties65);
            run47.Append(text65);

            paragraph75.Append(paragraphProperties37);
            paragraph75.Append(run47);

            var paragraph76 = new A.Paragraph();
            var paragraphProperties38 = new A.ParagraphProperties() {Level = 3};

            var run48 = new A.Run();

            var runProperties66 = new A.RunProperties() {Language = "en-US"};
            runProperties66.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text66 = new A.Text();
            text66.Text = "Fourth level";

            run48.Append(runProperties66);
            run48.Append(text66);

            paragraph76.Append(paragraphProperties38);
            paragraph76.Append(run48);

            var paragraph77 = new A.Paragraph();
            var paragraphProperties39 = new A.ParagraphProperties() {Level = 4};

            var run49 = new A.Run();

            var runProperties67 = new A.RunProperties() {Language = "en-US"};
            runProperties67.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text67 = new A.Text();
            text67.Text = "Fifth level";

            run49.Append(runProperties67);
            run49.Append(text67);
            var endParagraphRunProperties45 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph77.Append(paragraphProperties39);
            paragraph77.Append(run49);
            paragraph77.Append(endParagraphRunProperties45);

            textBody49.Append(bodyProperties49);
            textBody49.Append(listStyle49);
            textBody49.Append(paragraph73);
            textBody49.Append(paragraph74);
            textBody49.Append(paragraph75);
            textBody49.Append(paragraph76);
            textBody49.Append(paragraph77);

            shape49.Append(nonVisualShapeProperties49);
            shape49.Append(shapeProperties49);
            shape49.Append(textBody49);

            var shape50 = new Shape();

            var nonVisualShapeProperties50 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties61 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Date Placeholder 3"
            };

            var nonVisualShapeDrawingProperties50 = new NonVisualShapeDrawingProperties();
            var shapeLocks50 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties50.Append(shapeLocks50);

            var applicationNonVisualDrawingProperties61 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape50 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 10U
            };

            applicationNonVisualDrawingProperties61.Append(placeholderShape50);

            nonVisualShapeProperties50.Append(nonVisualDrawingProperties61);
            nonVisualShapeProperties50.Append(nonVisualShapeDrawingProperties50);
            nonVisualShapeProperties50.Append(applicationNonVisualDrawingProperties61);
            var shapeProperties50 = new ShapeProperties();

            var textBody50 = new TextBody();
            var bodyProperties50 = new A.BodyProperties();
            var listStyle50 = new A.ListStyle();

            var paragraph78 = new A.Paragraph();

            var field19 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties68 = new A.RunProperties() {Language = "en-GB"};
            runProperties68.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text68 = new A.Text();
            text68.Text = "04/06/2014";

            field19.Append(runProperties68);
            field19.Append(text68);
            var endParagraphRunProperties46 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph78.Append(field19);
            paragraph78.Append(endParagraphRunProperties46);

            textBody50.Append(bodyProperties50);
            textBody50.Append(listStyle50);
            textBody50.Append(paragraph78);

            shape50.Append(nonVisualShapeProperties50);
            shape50.Append(shapeProperties50);
            shape50.Append(textBody50);

            var shape51 = new Shape();

            var nonVisualShapeProperties51 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties62 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 5U,
                Name = "Footer Placeholder 4"
            };

            var nonVisualShapeDrawingProperties51 = new NonVisualShapeDrawingProperties();
            var shapeLocks51 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties51.Append(shapeLocks51);

            var applicationNonVisualDrawingProperties62 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape51 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 11U
            };

            applicationNonVisualDrawingProperties62.Append(placeholderShape51);

            nonVisualShapeProperties51.Append(nonVisualDrawingProperties62);
            nonVisualShapeProperties51.Append(nonVisualShapeDrawingProperties51);
            nonVisualShapeProperties51.Append(applicationNonVisualDrawingProperties62);
            var shapeProperties51 = new ShapeProperties();

            var textBody51 = new TextBody();
            var bodyProperties51 = new A.BodyProperties();
            var listStyle51 = new A.ListStyle();

            var paragraph79 = new A.Paragraph();
            var endParagraphRunProperties47 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph79.Append(endParagraphRunProperties47);

            textBody51.Append(bodyProperties51);
            textBody51.Append(listStyle51);
            textBody51.Append(paragraph79);

            shape51.Append(nonVisualShapeProperties51);
            shape51.Append(shapeProperties51);
            shape51.Append(textBody51);

            var shape52 = new Shape();

            var nonVisualShapeProperties52 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties63 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 6U,
                Name = "Slide Number Placeholder 5"
            };

            var nonVisualShapeDrawingProperties52 = new NonVisualShapeDrawingProperties();
            var shapeLocks52 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties52.Append(shapeLocks52);

            var applicationNonVisualDrawingProperties63 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape52 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 12U
            };

            applicationNonVisualDrawingProperties63.Append(placeholderShape52);

            nonVisualShapeProperties52.Append(nonVisualDrawingProperties63);
            nonVisualShapeProperties52.Append(nonVisualShapeDrawingProperties52);
            nonVisualShapeProperties52.Append(applicationNonVisualDrawingProperties63);
            var shapeProperties52 = new ShapeProperties();

            var textBody52 = new TextBody();
            var bodyProperties52 = new A.BodyProperties();
            var listStyle52 = new A.ListStyle();

            var paragraph80 = new A.Paragraph();

            var field20 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties69 = new A.RunProperties() {Language = "en-GB"};
            runProperties69.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text69 = new A.Text();
            text69.Text = "‹#›";

            field20.Append(runProperties69);
            field20.Append(text69);
            var endParagraphRunProperties48 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph80.Append(field20);
            paragraph80.Append(endParagraphRunProperties48);

            textBody52.Append(bodyProperties52);
            textBody52.Append(listStyle52);
            textBody52.Append(paragraph80);

            shape52.Append(nonVisualShapeProperties52);
            shape52.Append(shapeProperties52);
            shape52.Append(textBody52);

            shapeTree11.Append(nonVisualGroupShapeProperties11);
            shapeTree11.Append(groupShapeProperties11);
            shapeTree11.Append(shape48);
            shapeTree11.Append(shape49);
            shapeTree11.Append(shape50);
            shapeTree11.Append(shape51);
            shapeTree11.Append(shape52);

            var commonSlideDataExtensionList11 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension11 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId11 = new P14.CreationId() {Val = (UInt32Value) 3301124249U};
            creationId11.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension11.Append(creationId11);

            commonSlideDataExtensionList11.Append(commonSlideDataExtension11);

            commonSlideData11.Append(shapeTree11);
            commonSlideData11.Append(commonSlideDataExtensionList11);

            var colorMapOverride10 = new ColorMapOverride();
            var masterColorMapping10 = new A.MasterColorMapping();

            colorMapOverride10.Append(masterColorMapping10);

            slideLayout9.Append(commonSlideData11);
            slideLayout9.Append(colorMapOverride10);

            slideLayoutPart9.SlideLayout = slideLayout9;
        }

        // Generates content of slideLayoutPart10.
        private static void GenerateSlideLayoutPart10Content(SlideLayoutPart slideLayoutPart10)
        {
            var slideLayout10 = new SlideLayout() {Type = SlideLayoutValues.TwoObjects, Preserve = true};
            slideLayout10.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout10.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout10.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData12 = new CommonSlideData() {Name = "Two Content"};

            var shapeTree12 = new ShapeTree();

            var nonVisualGroupShapeProperties12 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties64 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties12 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties64 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties12.Append(nonVisualDrawingProperties64);
            nonVisualGroupShapeProperties12.Append(nonVisualGroupShapeDrawingProperties12);
            nonVisualGroupShapeProperties12.Append(applicationNonVisualDrawingProperties64);

            var groupShapeProperties12 = new GroupShapeProperties();

            var transformGroup12 = new A.TransformGroup();
            var offset30 = new A.Offset() {X = 0L, Y = 0L};
            var extents30 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset12 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents12 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup12.Append(offset30);
            transformGroup12.Append(extents30);
            transformGroup12.Append(childOffset12);
            transformGroup12.Append(childExtents12);

            groupShapeProperties12.Append(transformGroup12);

            var shape53 = new Shape();

            var nonVisualShapeProperties53 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties65 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 2U,
                Name = "Title 1"
            };

            var nonVisualShapeDrawingProperties53 = new NonVisualShapeDrawingProperties();
            var shapeLocks53 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties53.Append(shapeLocks53);

            var applicationNonVisualDrawingProperties65 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape53 = new PlaceholderShape() {Type = PlaceholderValues.Title};

            applicationNonVisualDrawingProperties65.Append(placeholderShape53);

            nonVisualShapeProperties53.Append(nonVisualDrawingProperties65);
            nonVisualShapeProperties53.Append(nonVisualShapeDrawingProperties53);
            nonVisualShapeProperties53.Append(applicationNonVisualDrawingProperties65);
            var shapeProperties53 = new ShapeProperties();

            var textBody53 = new TextBody();
            var bodyProperties53 = new A.BodyProperties();
            var listStyle53 = new A.ListStyle();

            var paragraph81 = new A.Paragraph();

            var run50 = new A.Run();

            var runProperties70 = new A.RunProperties() {Language = "en-US"};
            runProperties70.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text70 = new A.Text();
            text70.Text = "Click to edit Master title style";

            run50.Append(runProperties70);
            run50.Append(text70);
            var endParagraphRunProperties49 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph81.Append(run50);
            paragraph81.Append(endParagraphRunProperties49);

            textBody53.Append(bodyProperties53);
            textBody53.Append(listStyle53);
            textBody53.Append(paragraph81);

            shape53.Append(nonVisualShapeProperties53);
            shape53.Append(shapeProperties53);
            shape53.Append(textBody53);

            var shape54 = new Shape();

            var nonVisualShapeProperties54 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties66 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Content Placeholder 2"
            };

            var nonVisualShapeDrawingProperties54 = new NonVisualShapeDrawingProperties();
            var shapeLocks54 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties54.Append(shapeLocks54);

            var applicationNonVisualDrawingProperties66 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape54 = new PlaceholderShape() {
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 1U
            };

            applicationNonVisualDrawingProperties66.Append(placeholderShape54);

            nonVisualShapeProperties54.Append(nonVisualDrawingProperties66);
            nonVisualShapeProperties54.Append(nonVisualShapeDrawingProperties54);
            nonVisualShapeProperties54.Append(applicationNonVisualDrawingProperties66);

            var shapeProperties54 = new ShapeProperties();

            var transform2D19 = new A.Transform2D();
            var offset31 = new A.Offset() {X = 457200L, Y = 1600200L};
            var extents31 = new A.Extents() {Cx = 4038600L, Cy = 4525963L};

            transform2D19.Append(offset31);
            transform2D19.Append(extents31);

            shapeProperties54.Append(transform2D19);

            var textBody54 = new TextBody();
            var bodyProperties54 = new A.BodyProperties();

            var listStyle54 = new A.ListStyle();

            var level1ParagraphProperties19 = new A.Level1ParagraphProperties();
            var defaultRunProperties109 = new A.DefaultRunProperties() {FontSize = 2800};

            level1ParagraphProperties19.Append(defaultRunProperties109);

            var level2ParagraphProperties12 = new A.Level2ParagraphProperties();
            var defaultRunProperties110 = new A.DefaultRunProperties() {FontSize = 2400};

            level2ParagraphProperties12.Append(defaultRunProperties110);

            var level3ParagraphProperties12 = new A.Level3ParagraphProperties();
            var defaultRunProperties111 = new A.DefaultRunProperties() {FontSize = 2000};

            level3ParagraphProperties12.Append(defaultRunProperties111);

            var level4ParagraphProperties12 = new A.Level4ParagraphProperties();
            var defaultRunProperties112 = new A.DefaultRunProperties() {FontSize = 1800};

            level4ParagraphProperties12.Append(defaultRunProperties112);

            var level5ParagraphProperties12 = new A.Level5ParagraphProperties();
            var defaultRunProperties113 = new A.DefaultRunProperties() {FontSize = 1800};

            level5ParagraphProperties12.Append(defaultRunProperties113);

            var level6ParagraphProperties12 = new A.Level6ParagraphProperties();
            var defaultRunProperties114 = new A.DefaultRunProperties() {FontSize = 1800};

            level6ParagraphProperties12.Append(defaultRunProperties114);

            var level7ParagraphProperties12 = new A.Level7ParagraphProperties();
            var defaultRunProperties115 = new A.DefaultRunProperties() {FontSize = 1800};

            level7ParagraphProperties12.Append(defaultRunProperties115);

            var level8ParagraphProperties12 = new A.Level8ParagraphProperties();
            var defaultRunProperties116 = new A.DefaultRunProperties() {FontSize = 1800};

            level8ParagraphProperties12.Append(defaultRunProperties116);

            var level9ParagraphProperties12 = new A.Level9ParagraphProperties();
            var defaultRunProperties117 = new A.DefaultRunProperties() {FontSize = 1800};

            level9ParagraphProperties12.Append(defaultRunProperties117);

            listStyle54.Append(level1ParagraphProperties19);
            listStyle54.Append(level2ParagraphProperties12);
            listStyle54.Append(level3ParagraphProperties12);
            listStyle54.Append(level4ParagraphProperties12);
            listStyle54.Append(level5ParagraphProperties12);
            listStyle54.Append(level6ParagraphProperties12);
            listStyle54.Append(level7ParagraphProperties12);
            listStyle54.Append(level8ParagraphProperties12);
            listStyle54.Append(level9ParagraphProperties12);

            var paragraph82 = new A.Paragraph();
            var paragraphProperties40 = new A.ParagraphProperties() {Level = 0};

            var run51 = new A.Run();

            var runProperties71 = new A.RunProperties() {Language = "en-US"};
            runProperties71.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text71 = new A.Text();
            text71.Text = "Click to edit Master text styles";

            run51.Append(runProperties71);
            run51.Append(text71);

            paragraph82.Append(paragraphProperties40);
            paragraph82.Append(run51);

            var paragraph83 = new A.Paragraph();
            var paragraphProperties41 = new A.ParagraphProperties() {Level = 1};

            var run52 = new A.Run();

            var runProperties72 = new A.RunProperties() {Language = "en-US"};
            runProperties72.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text72 = new A.Text();
            text72.Text = "Second level";

            run52.Append(runProperties72);
            run52.Append(text72);

            paragraph83.Append(paragraphProperties41);
            paragraph83.Append(run52);

            var paragraph84 = new A.Paragraph();
            var paragraphProperties42 = new A.ParagraphProperties() {Level = 2};

            var run53 = new A.Run();

            var runProperties73 = new A.RunProperties() {Language = "en-US"};
            runProperties73.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text73 = new A.Text();
            text73.Text = "Third level";

            run53.Append(runProperties73);
            run53.Append(text73);

            paragraph84.Append(paragraphProperties42);
            paragraph84.Append(run53);

            var paragraph85 = new A.Paragraph();
            var paragraphProperties43 = new A.ParagraphProperties() {Level = 3};

            var run54 = new A.Run();

            var runProperties74 = new A.RunProperties() {Language = "en-US"};
            runProperties74.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text74 = new A.Text();
            text74.Text = "Fourth level";

            run54.Append(runProperties74);
            run54.Append(text74);

            paragraph85.Append(paragraphProperties43);
            paragraph85.Append(run54);

            var paragraph86 = new A.Paragraph();
            var paragraphProperties44 = new A.ParagraphProperties() {Level = 4};

            var run55 = new A.Run();

            var runProperties75 = new A.RunProperties() {Language = "en-US"};
            runProperties75.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text75 = new A.Text();
            text75.Text = "Fifth level";

            run55.Append(runProperties75);
            run55.Append(text75);
            var endParagraphRunProperties50 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph86.Append(paragraphProperties44);
            paragraph86.Append(run55);
            paragraph86.Append(endParagraphRunProperties50);

            textBody54.Append(bodyProperties54);
            textBody54.Append(listStyle54);
            textBody54.Append(paragraph82);
            textBody54.Append(paragraph83);
            textBody54.Append(paragraph84);
            textBody54.Append(paragraph85);
            textBody54.Append(paragraph86);

            shape54.Append(nonVisualShapeProperties54);
            shape54.Append(shapeProperties54);
            shape54.Append(textBody54);

            var shape55 = new Shape();

            var nonVisualShapeProperties55 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties67 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Content Placeholder 3"
            };

            var nonVisualShapeDrawingProperties55 = new NonVisualShapeDrawingProperties();
            var shapeLocks55 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties55.Append(shapeLocks55);

            var applicationNonVisualDrawingProperties67 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape55 = new PlaceholderShape() {
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 2U
            };

            applicationNonVisualDrawingProperties67.Append(placeholderShape55);

            nonVisualShapeProperties55.Append(nonVisualDrawingProperties67);
            nonVisualShapeProperties55.Append(nonVisualShapeDrawingProperties55);
            nonVisualShapeProperties55.Append(applicationNonVisualDrawingProperties67);

            var shapeProperties55 = new ShapeProperties();

            var transform2D20 = new A.Transform2D();
            var offset32 = new A.Offset() {X = 4648200L, Y = 1600200L};
            var extents32 = new A.Extents() {Cx = 4038600L, Cy = 4525963L};

            transform2D20.Append(offset32);
            transform2D20.Append(extents32);

            shapeProperties55.Append(transform2D20);

            var textBody55 = new TextBody();
            var bodyProperties55 = new A.BodyProperties();

            var listStyle55 = new A.ListStyle();

            var level1ParagraphProperties20 = new A.Level1ParagraphProperties();
            var defaultRunProperties118 = new A.DefaultRunProperties() {FontSize = 2800};

            level1ParagraphProperties20.Append(defaultRunProperties118);

            var level2ParagraphProperties13 = new A.Level2ParagraphProperties();
            var defaultRunProperties119 = new A.DefaultRunProperties() {FontSize = 2400};

            level2ParagraphProperties13.Append(defaultRunProperties119);

            var level3ParagraphProperties13 = new A.Level3ParagraphProperties();
            var defaultRunProperties120 = new A.DefaultRunProperties() {FontSize = 2000};

            level3ParagraphProperties13.Append(defaultRunProperties120);

            var level4ParagraphProperties13 = new A.Level4ParagraphProperties();
            var defaultRunProperties121 = new A.DefaultRunProperties() {FontSize = 1800};

            level4ParagraphProperties13.Append(defaultRunProperties121);

            var level5ParagraphProperties13 = new A.Level5ParagraphProperties();
            var defaultRunProperties122 = new A.DefaultRunProperties() {FontSize = 1800};

            level5ParagraphProperties13.Append(defaultRunProperties122);

            var level6ParagraphProperties13 = new A.Level6ParagraphProperties();
            var defaultRunProperties123 = new A.DefaultRunProperties() {FontSize = 1800};

            level6ParagraphProperties13.Append(defaultRunProperties123);

            var level7ParagraphProperties13 = new A.Level7ParagraphProperties();
            var defaultRunProperties124 = new A.DefaultRunProperties() {FontSize = 1800};

            level7ParagraphProperties13.Append(defaultRunProperties124);

            var level8ParagraphProperties13 = new A.Level8ParagraphProperties();
            var defaultRunProperties125 = new A.DefaultRunProperties() {FontSize = 1800};

            level8ParagraphProperties13.Append(defaultRunProperties125);

            var level9ParagraphProperties13 = new A.Level9ParagraphProperties();
            var defaultRunProperties126 = new A.DefaultRunProperties() {FontSize = 1800};

            level9ParagraphProperties13.Append(defaultRunProperties126);

            listStyle55.Append(level1ParagraphProperties20);
            listStyle55.Append(level2ParagraphProperties13);
            listStyle55.Append(level3ParagraphProperties13);
            listStyle55.Append(level4ParagraphProperties13);
            listStyle55.Append(level5ParagraphProperties13);
            listStyle55.Append(level6ParagraphProperties13);
            listStyle55.Append(level7ParagraphProperties13);
            listStyle55.Append(level8ParagraphProperties13);
            listStyle55.Append(level9ParagraphProperties13);

            var paragraph87 = new A.Paragraph();
            var paragraphProperties45 = new A.ParagraphProperties() {Level = 0};

            var run56 = new A.Run();

            var runProperties76 = new A.RunProperties() {Language = "en-US"};
            runProperties76.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text76 = new A.Text();
            text76.Text = "Click to edit Master text styles";

            run56.Append(runProperties76);
            run56.Append(text76);

            paragraph87.Append(paragraphProperties45);
            paragraph87.Append(run56);

            var paragraph88 = new A.Paragraph();
            var paragraphProperties46 = new A.ParagraphProperties() {Level = 1};

            var run57 = new A.Run();

            var runProperties77 = new A.RunProperties() {Language = "en-US"};
            runProperties77.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text77 = new A.Text();
            text77.Text = "Second level";

            run57.Append(runProperties77);
            run57.Append(text77);

            paragraph88.Append(paragraphProperties46);
            paragraph88.Append(run57);

            var paragraph89 = new A.Paragraph();
            var paragraphProperties47 = new A.ParagraphProperties() {Level = 2};

            var run58 = new A.Run();

            var runProperties78 = new A.RunProperties() {Language = "en-US"};
            runProperties78.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text78 = new A.Text();
            text78.Text = "Third level";

            run58.Append(runProperties78);
            run58.Append(text78);

            paragraph89.Append(paragraphProperties47);
            paragraph89.Append(run58);

            var paragraph90 = new A.Paragraph();
            var paragraphProperties48 = new A.ParagraphProperties() {Level = 3};

            var run59 = new A.Run();

            var runProperties79 = new A.RunProperties() {Language = "en-US"};
            runProperties79.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text79 = new A.Text();
            text79.Text = "Fourth level";

            run59.Append(runProperties79);
            run59.Append(text79);

            paragraph90.Append(paragraphProperties48);
            paragraph90.Append(run59);

            var paragraph91 = new A.Paragraph();
            var paragraphProperties49 = new A.ParagraphProperties() {Level = 4};

            var run60 = new A.Run();

            var runProperties80 = new A.RunProperties() {Language = "en-US"};
            runProperties80.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text80 = new A.Text();
            text80.Text = "Fifth level";

            run60.Append(runProperties80);
            run60.Append(text80);
            var endParagraphRunProperties51 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph91.Append(paragraphProperties49);
            paragraph91.Append(run60);
            paragraph91.Append(endParagraphRunProperties51);

            textBody55.Append(bodyProperties55);
            textBody55.Append(listStyle55);
            textBody55.Append(paragraph87);
            textBody55.Append(paragraph88);
            textBody55.Append(paragraph89);
            textBody55.Append(paragraph90);
            textBody55.Append(paragraph91);

            shape55.Append(nonVisualShapeProperties55);
            shape55.Append(shapeProperties55);
            shape55.Append(textBody55);

            var shape56 = new Shape();

            var nonVisualShapeProperties56 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties68 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 5U,
                Name = "Date Placeholder 4"
            };

            var nonVisualShapeDrawingProperties56 = new NonVisualShapeDrawingProperties();
            var shapeLocks56 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties56.Append(shapeLocks56);

            var applicationNonVisualDrawingProperties68 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape56 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 10U
            };

            applicationNonVisualDrawingProperties68.Append(placeholderShape56);

            nonVisualShapeProperties56.Append(nonVisualDrawingProperties68);
            nonVisualShapeProperties56.Append(nonVisualShapeDrawingProperties56);
            nonVisualShapeProperties56.Append(applicationNonVisualDrawingProperties68);
            var shapeProperties56 = new ShapeProperties();

            var textBody56 = new TextBody();
            var bodyProperties56 = new A.BodyProperties();
            var listStyle56 = new A.ListStyle();

            var paragraph92 = new A.Paragraph();

            var field21 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties81 = new A.RunProperties() {Language = "en-GB"};
            runProperties81.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text81 = new A.Text();
            text81.Text = "04/06/2014";

            field21.Append(runProperties81);
            field21.Append(text81);
            var endParagraphRunProperties52 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph92.Append(field21);
            paragraph92.Append(endParagraphRunProperties52);

            textBody56.Append(bodyProperties56);
            textBody56.Append(listStyle56);
            textBody56.Append(paragraph92);

            shape56.Append(nonVisualShapeProperties56);
            shape56.Append(shapeProperties56);
            shape56.Append(textBody56);

            var shape57 = new Shape();

            var nonVisualShapeProperties57 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties69 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 6U,
                Name = "Footer Placeholder 5"
            };

            var nonVisualShapeDrawingProperties57 = new NonVisualShapeDrawingProperties();
            var shapeLocks57 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties57.Append(shapeLocks57);

            var applicationNonVisualDrawingProperties69 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape57 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 11U
            };

            applicationNonVisualDrawingProperties69.Append(placeholderShape57);

            nonVisualShapeProperties57.Append(nonVisualDrawingProperties69);
            nonVisualShapeProperties57.Append(nonVisualShapeDrawingProperties57);
            nonVisualShapeProperties57.Append(applicationNonVisualDrawingProperties69);
            var shapeProperties57 = new ShapeProperties();

            var textBody57 = new TextBody();
            var bodyProperties57 = new A.BodyProperties();
            var listStyle57 = new A.ListStyle();

            var paragraph93 = new A.Paragraph();
            var endParagraphRunProperties53 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph93.Append(endParagraphRunProperties53);

            textBody57.Append(bodyProperties57);
            textBody57.Append(listStyle57);
            textBody57.Append(paragraph93);

            shape57.Append(nonVisualShapeProperties57);
            shape57.Append(shapeProperties57);
            shape57.Append(textBody57);

            var shape58 = new Shape();

            var nonVisualShapeProperties58 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties70 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 7U,
                Name = "Slide Number Placeholder 6"
            };

            var nonVisualShapeDrawingProperties58 = new NonVisualShapeDrawingProperties();
            var shapeLocks58 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties58.Append(shapeLocks58);

            var applicationNonVisualDrawingProperties70 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape58 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 12U
            };

            applicationNonVisualDrawingProperties70.Append(placeholderShape58);

            nonVisualShapeProperties58.Append(nonVisualDrawingProperties70);
            nonVisualShapeProperties58.Append(nonVisualShapeDrawingProperties58);
            nonVisualShapeProperties58.Append(applicationNonVisualDrawingProperties70);
            var shapeProperties58 = new ShapeProperties();

            var textBody58 = new TextBody();
            var bodyProperties58 = new A.BodyProperties();
            var listStyle58 = new A.ListStyle();

            var paragraph94 = new A.Paragraph();

            var field22 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties82 = new A.RunProperties() {Language = "en-GB"};
            runProperties82.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text82 = new A.Text();
            text82.Text = "‹#›";

            field22.Append(runProperties82);
            field22.Append(text82);
            var endParagraphRunProperties54 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph94.Append(field22);
            paragraph94.Append(endParagraphRunProperties54);

            textBody58.Append(bodyProperties58);
            textBody58.Append(listStyle58);
            textBody58.Append(paragraph94);

            shape58.Append(nonVisualShapeProperties58);
            shape58.Append(shapeProperties58);
            shape58.Append(textBody58);

            shapeTree12.Append(nonVisualGroupShapeProperties12);
            shapeTree12.Append(groupShapeProperties12);
            shapeTree12.Append(shape53);
            shapeTree12.Append(shape54);
            shapeTree12.Append(shape55);
            shapeTree12.Append(shape56);
            shapeTree12.Append(shape57);
            shapeTree12.Append(shape58);

            var commonSlideDataExtensionList12 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension12 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId12 = new P14.CreationId() {Val = (UInt32Value) 2622313794U};
            creationId12.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension12.Append(creationId12);

            commonSlideDataExtensionList12.Append(commonSlideDataExtension12);

            commonSlideData12.Append(shapeTree12);
            commonSlideData12.Append(commonSlideDataExtensionList12);

            var colorMapOverride11 = new ColorMapOverride();
            var masterColorMapping11 = new A.MasterColorMapping();

            colorMapOverride11.Append(masterColorMapping11);

            slideLayout10.Append(commonSlideData12);
            slideLayout10.Append(colorMapOverride11);

            slideLayoutPart10.SlideLayout = slideLayout10;
        }

        // Generates content of slideLayoutPart11.
        private static void GenerateSlideLayoutPart11Content(SlideLayoutPart slideLayoutPart11)
        {
            var slideLayout11 = new SlideLayout() {Type = SlideLayoutValues.PictureText, Preserve = true};
            slideLayout11.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout11.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout11.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var commonSlideData13 = new CommonSlideData() {Name = "Picture with Caption"};

            var shapeTree13 = new ShapeTree();

            var nonVisualGroupShapeProperties13 = new NonVisualGroupShapeProperties();
            var nonVisualDrawingProperties71 = new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""};
            var nonVisualGroupShapeDrawingProperties13 = new NonVisualGroupShapeDrawingProperties();
            var applicationNonVisualDrawingProperties71 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties13.Append(nonVisualDrawingProperties71);
            nonVisualGroupShapeProperties13.Append(nonVisualGroupShapeDrawingProperties13);
            nonVisualGroupShapeProperties13.Append(applicationNonVisualDrawingProperties71);

            var groupShapeProperties13 = new GroupShapeProperties();

            var transformGroup13 = new A.TransformGroup();
            var offset33 = new A.Offset() {X = 0L, Y = 0L};
            var extents33 = new A.Extents() {Cx = 0L, Cy = 0L};
            var childOffset13 = new A.ChildOffset() {X = 0L, Y = 0L};
            var childExtents13 = new A.ChildExtents() {Cx = 0L, Cy = 0L};

            transformGroup13.Append(offset33);
            transformGroup13.Append(extents33);
            transformGroup13.Append(childOffset13);
            transformGroup13.Append(childExtents13);

            groupShapeProperties13.Append(transformGroup13);

            var shape59 = new Shape();

            var nonVisualShapeProperties59 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties72 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 2U,
                Name = "Title 1"
            };

            var nonVisualShapeDrawingProperties59 = new NonVisualShapeDrawingProperties();
            var shapeLocks59 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties59.Append(shapeLocks59);

            var applicationNonVisualDrawingProperties72 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape59 = new PlaceholderShape() {Type = PlaceholderValues.Title};

            applicationNonVisualDrawingProperties72.Append(placeholderShape59);

            nonVisualShapeProperties59.Append(nonVisualDrawingProperties72);
            nonVisualShapeProperties59.Append(nonVisualShapeDrawingProperties59);
            nonVisualShapeProperties59.Append(applicationNonVisualDrawingProperties72);

            var shapeProperties59 = new ShapeProperties();

            var transform2D21 = new A.Transform2D();
            var offset34 = new A.Offset() {X = 1792288L, Y = 4800600L};
            var extents34 = new A.Extents() {Cx = 5486400L, Cy = 566738L};

            transform2D21.Append(offset34);
            transform2D21.Append(extents34);

            shapeProperties59.Append(transform2D21);

            var textBody59 = new TextBody();
            var bodyProperties59 = new A.BodyProperties() {Anchor = A.TextAnchoringTypeValues.Bottom};

            var listStyle59 = new A.ListStyle();

            var level1ParagraphProperties21 = new A.Level1ParagraphProperties() {
                Alignment = A.TextAlignmentTypeValues.Left
            };
            var defaultRunProperties127 = new A.DefaultRunProperties() {FontSize = 2000, Bold = true};

            level1ParagraphProperties21.Append(defaultRunProperties127);

            listStyle59.Append(level1ParagraphProperties21);

            var paragraph95 = new A.Paragraph();

            var run61 = new A.Run();

            var runProperties83 = new A.RunProperties() {Language = "en-US"};
            runProperties83.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text83 = new A.Text();
            text83.Text = "Click to edit Master title style";

            run61.Append(runProperties83);
            run61.Append(text83);
            var endParagraphRunProperties55 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph95.Append(run61);
            paragraph95.Append(endParagraphRunProperties55);

            textBody59.Append(bodyProperties59);
            textBody59.Append(listStyle59);
            textBody59.Append(paragraph95);

            shape59.Append(nonVisualShapeProperties59);
            shape59.Append(shapeProperties59);
            shape59.Append(textBody59);

            var shape60 = new Shape();

            var nonVisualShapeProperties60 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties73 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 3U,
                Name = "Picture Placeholder 2"
            };

            var nonVisualShapeDrawingProperties60 = new NonVisualShapeDrawingProperties();
            var shapeLocks60 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties60.Append(shapeLocks60);

            var applicationNonVisualDrawingProperties73 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape60 = new PlaceholderShape() {Type = PlaceholderValues.Picture, Index = (UInt32Value) 1U};

            applicationNonVisualDrawingProperties73.Append(placeholderShape60);

            nonVisualShapeProperties60.Append(nonVisualDrawingProperties73);
            nonVisualShapeProperties60.Append(nonVisualShapeDrawingProperties60);
            nonVisualShapeProperties60.Append(applicationNonVisualDrawingProperties73);

            var shapeProperties60 = new ShapeProperties();

            var transform2D22 = new A.Transform2D();
            var offset35 = new A.Offset() {X = 1792288L, Y = 612775L};
            var extents35 = new A.Extents() {Cx = 5486400L, Cy = 4114800L};

            transform2D22.Append(offset35);
            transform2D22.Append(extents35);

            shapeProperties60.Append(transform2D22);

            var textBody60 = new TextBody();
            var bodyProperties60 = new A.BodyProperties();

            var listStyle60 = new A.ListStyle();

            var level1ParagraphProperties22 = new A.Level1ParagraphProperties() {LeftMargin = 0, Indent = 0};
            var noBullet47 = new A.NoBullet();
            var defaultRunProperties128 = new A.DefaultRunProperties() {FontSize = 3200};

            level1ParagraphProperties22.Append(noBullet47);
            level1ParagraphProperties22.Append(defaultRunProperties128);

            var level2ParagraphProperties14 = new A.Level2ParagraphProperties() {LeftMargin = 457200, Indent = 0};
            var noBullet48 = new A.NoBullet();
            var defaultRunProperties129 = new A.DefaultRunProperties() {FontSize = 2800};

            level2ParagraphProperties14.Append(noBullet48);
            level2ParagraphProperties14.Append(defaultRunProperties129);

            var level3ParagraphProperties14 = new A.Level3ParagraphProperties() {LeftMargin = 914400, Indent = 0};
            var noBullet49 = new A.NoBullet();
            var defaultRunProperties130 = new A.DefaultRunProperties() {FontSize = 2400};

            level3ParagraphProperties14.Append(noBullet49);
            level3ParagraphProperties14.Append(defaultRunProperties130);

            var level4ParagraphProperties14 = new A.Level4ParagraphProperties() {LeftMargin = 1371600, Indent = 0};
            var noBullet50 = new A.NoBullet();
            var defaultRunProperties131 = new A.DefaultRunProperties() {FontSize = 2000};

            level4ParagraphProperties14.Append(noBullet50);
            level4ParagraphProperties14.Append(defaultRunProperties131);

            var level5ParagraphProperties14 = new A.Level5ParagraphProperties() {LeftMargin = 1828800, Indent = 0};
            var noBullet51 = new A.NoBullet();
            var defaultRunProperties132 = new A.DefaultRunProperties() {FontSize = 2000};

            level5ParagraphProperties14.Append(noBullet51);
            level5ParagraphProperties14.Append(defaultRunProperties132);

            var level6ParagraphProperties14 = new A.Level6ParagraphProperties() {LeftMargin = 2286000, Indent = 0};
            var noBullet52 = new A.NoBullet();
            var defaultRunProperties133 = new A.DefaultRunProperties() {FontSize = 2000};

            level6ParagraphProperties14.Append(noBullet52);
            level6ParagraphProperties14.Append(defaultRunProperties133);

            var level7ParagraphProperties14 = new A.Level7ParagraphProperties() {LeftMargin = 2743200, Indent = 0};
            var noBullet53 = new A.NoBullet();
            var defaultRunProperties134 = new A.DefaultRunProperties() {FontSize = 2000};

            level7ParagraphProperties14.Append(noBullet53);
            level7ParagraphProperties14.Append(defaultRunProperties134);

            var level8ParagraphProperties14 = new A.Level8ParagraphProperties() {LeftMargin = 3200400, Indent = 0};
            var noBullet54 = new A.NoBullet();
            var defaultRunProperties135 = new A.DefaultRunProperties() {FontSize = 2000};

            level8ParagraphProperties14.Append(noBullet54);
            level8ParagraphProperties14.Append(defaultRunProperties135);

            var level9ParagraphProperties14 = new A.Level9ParagraphProperties() {LeftMargin = 3657600, Indent = 0};
            var noBullet55 = new A.NoBullet();
            var defaultRunProperties136 = new A.DefaultRunProperties() {FontSize = 2000};

            level9ParagraphProperties14.Append(noBullet55);
            level9ParagraphProperties14.Append(defaultRunProperties136);

            listStyle60.Append(level1ParagraphProperties22);
            listStyle60.Append(level2ParagraphProperties14);
            listStyle60.Append(level3ParagraphProperties14);
            listStyle60.Append(level4ParagraphProperties14);
            listStyle60.Append(level5ParagraphProperties14);
            listStyle60.Append(level6ParagraphProperties14);
            listStyle60.Append(level7ParagraphProperties14);
            listStyle60.Append(level8ParagraphProperties14);
            listStyle60.Append(level9ParagraphProperties14);

            var paragraph96 = new A.Paragraph();
            var endParagraphRunProperties56 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph96.Append(endParagraphRunProperties56);

            textBody60.Append(bodyProperties60);
            textBody60.Append(listStyle60);
            textBody60.Append(paragraph96);

            shape60.Append(nonVisualShapeProperties60);
            shape60.Append(shapeProperties60);
            shape60.Append(textBody60);

            var shape61 = new Shape();

            var nonVisualShapeProperties61 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties74 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 4U,
                Name = "Text Placeholder 3"
            };

            var nonVisualShapeDrawingProperties61 = new NonVisualShapeDrawingProperties();
            var shapeLocks61 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties61.Append(shapeLocks61);

            var applicationNonVisualDrawingProperties74 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape61 = new PlaceholderShape() {
                Type = PlaceholderValues.Body,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 2U
            };

            applicationNonVisualDrawingProperties74.Append(placeholderShape61);

            nonVisualShapeProperties61.Append(nonVisualDrawingProperties74);
            nonVisualShapeProperties61.Append(nonVisualShapeDrawingProperties61);
            nonVisualShapeProperties61.Append(applicationNonVisualDrawingProperties74);

            var shapeProperties61 = new ShapeProperties();

            var transform2D23 = new A.Transform2D();
            var offset36 = new A.Offset() {X = 1792288L, Y = 5367338L};
            var extents36 = new A.Extents() {Cx = 5486400L, Cy = 804862L};

            transform2D23.Append(offset36);
            transform2D23.Append(extents36);

            shapeProperties61.Append(transform2D23);

            var textBody61 = new TextBody();
            var bodyProperties61 = new A.BodyProperties();

            var listStyle61 = new A.ListStyle();

            var level1ParagraphProperties23 = new A.Level1ParagraphProperties() {LeftMargin = 0, Indent = 0};
            var noBullet56 = new A.NoBullet();
            var defaultRunProperties137 = new A.DefaultRunProperties() {FontSize = 1400};

            level1ParagraphProperties23.Append(noBullet56);
            level1ParagraphProperties23.Append(defaultRunProperties137);

            var level2ParagraphProperties15 = new A.Level2ParagraphProperties() {LeftMargin = 457200, Indent = 0};
            var noBullet57 = new A.NoBullet();
            var defaultRunProperties138 = new A.DefaultRunProperties() {FontSize = 1200};

            level2ParagraphProperties15.Append(noBullet57);
            level2ParagraphProperties15.Append(defaultRunProperties138);

            var level3ParagraphProperties15 = new A.Level3ParagraphProperties() {LeftMargin = 914400, Indent = 0};
            var noBullet58 = new A.NoBullet();
            var defaultRunProperties139 = new A.DefaultRunProperties() {FontSize = 1000};

            level3ParagraphProperties15.Append(noBullet58);
            level3ParagraphProperties15.Append(defaultRunProperties139);

            var level4ParagraphProperties15 = new A.Level4ParagraphProperties() {LeftMargin = 1371600, Indent = 0};
            var noBullet59 = new A.NoBullet();
            var defaultRunProperties140 = new A.DefaultRunProperties() {FontSize = 900};

            level4ParagraphProperties15.Append(noBullet59);
            level4ParagraphProperties15.Append(defaultRunProperties140);

            var level5ParagraphProperties15 = new A.Level5ParagraphProperties() {LeftMargin = 1828800, Indent = 0};
            var noBullet60 = new A.NoBullet();
            var defaultRunProperties141 = new A.DefaultRunProperties() {FontSize = 900};

            level5ParagraphProperties15.Append(noBullet60);
            level5ParagraphProperties15.Append(defaultRunProperties141);

            var level6ParagraphProperties15 = new A.Level6ParagraphProperties() {LeftMargin = 2286000, Indent = 0};
            var noBullet61 = new A.NoBullet();
            var defaultRunProperties142 = new A.DefaultRunProperties() {FontSize = 900};

            level6ParagraphProperties15.Append(noBullet61);
            level6ParagraphProperties15.Append(defaultRunProperties142);

            var level7ParagraphProperties15 = new A.Level7ParagraphProperties() {LeftMargin = 2743200, Indent = 0};
            var noBullet62 = new A.NoBullet();
            var defaultRunProperties143 = new A.DefaultRunProperties() {FontSize = 900};

            level7ParagraphProperties15.Append(noBullet62);
            level7ParagraphProperties15.Append(defaultRunProperties143);

            var level8ParagraphProperties15 = new A.Level8ParagraphProperties() {LeftMargin = 3200400, Indent = 0};
            var noBullet63 = new A.NoBullet();
            var defaultRunProperties144 = new A.DefaultRunProperties() {FontSize = 900};

            level8ParagraphProperties15.Append(noBullet63);
            level8ParagraphProperties15.Append(defaultRunProperties144);

            var level9ParagraphProperties15 = new A.Level9ParagraphProperties() {LeftMargin = 3657600, Indent = 0};
            var noBullet64 = new A.NoBullet();
            var defaultRunProperties145 = new A.DefaultRunProperties() {FontSize = 900};

            level9ParagraphProperties15.Append(noBullet64);
            level9ParagraphProperties15.Append(defaultRunProperties145);

            listStyle61.Append(level1ParagraphProperties23);
            listStyle61.Append(level2ParagraphProperties15);
            listStyle61.Append(level3ParagraphProperties15);
            listStyle61.Append(level4ParagraphProperties15);
            listStyle61.Append(level5ParagraphProperties15);
            listStyle61.Append(level6ParagraphProperties15);
            listStyle61.Append(level7ParagraphProperties15);
            listStyle61.Append(level8ParagraphProperties15);
            listStyle61.Append(level9ParagraphProperties15);

            var paragraph97 = new A.Paragraph();
            var paragraphProperties50 = new A.ParagraphProperties() {Level = 0};

            var run62 = new A.Run();

            var runProperties84 = new A.RunProperties() {Language = "en-US"};
            runProperties84.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text84 = new A.Text();
            text84.Text = "Click to edit Master text styles";

            run62.Append(runProperties84);
            run62.Append(text84);

            paragraph97.Append(paragraphProperties50);
            paragraph97.Append(run62);

            textBody61.Append(bodyProperties61);
            textBody61.Append(listStyle61);
            textBody61.Append(paragraph97);

            shape61.Append(nonVisualShapeProperties61);
            shape61.Append(shapeProperties61);
            shape61.Append(textBody61);

            var shape62 = new Shape();

            var nonVisualShapeProperties62 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties75 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 5U,
                Name = "Date Placeholder 4"
            };

            var nonVisualShapeDrawingProperties62 = new NonVisualShapeDrawingProperties();
            var shapeLocks62 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties62.Append(shapeLocks62);

            var applicationNonVisualDrawingProperties75 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape62 = new PlaceholderShape() {
                Type = PlaceholderValues.DateAndTime,
                Size = PlaceholderSizeValues.Half,
                Index = (UInt32Value) 10U
            };

            applicationNonVisualDrawingProperties75.Append(placeholderShape62);

            nonVisualShapeProperties62.Append(nonVisualDrawingProperties75);
            nonVisualShapeProperties62.Append(nonVisualShapeDrawingProperties62);
            nonVisualShapeProperties62.Append(applicationNonVisualDrawingProperties75);
            var shapeProperties62 = new ShapeProperties();

            var textBody62 = new TextBody();
            var bodyProperties62 = new A.BodyProperties();
            var listStyle62 = new A.ListStyle();

            var paragraph98 = new A.Paragraph();

            var field23 = new A.Field() {Id = "{32194167-FB73-4A8E-B54A-1D41D9EB4FBF}", Type = "datetimeFigureOut"};

            var runProperties85 = new A.RunProperties() {Language = "en-GB"};
            runProperties85.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text85 = new A.Text();
            text85.Text = "04/06/2014";

            field23.Append(runProperties85);
            field23.Append(text85);
            var endParagraphRunProperties57 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph98.Append(field23);
            paragraph98.Append(endParagraphRunProperties57);

            textBody62.Append(bodyProperties62);
            textBody62.Append(listStyle62);
            textBody62.Append(paragraph98);

            shape62.Append(nonVisualShapeProperties62);
            shape62.Append(shapeProperties62);
            shape62.Append(textBody62);

            var shape63 = new Shape();

            var nonVisualShapeProperties63 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties76 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 6U,
                Name = "Footer Placeholder 5"
            };

            var nonVisualShapeDrawingProperties63 = new NonVisualShapeDrawingProperties();
            var shapeLocks63 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties63.Append(shapeLocks63);

            var applicationNonVisualDrawingProperties76 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape63 = new PlaceholderShape() {
                Type = PlaceholderValues.Footer,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 11U
            };

            applicationNonVisualDrawingProperties76.Append(placeholderShape63);

            nonVisualShapeProperties63.Append(nonVisualDrawingProperties76);
            nonVisualShapeProperties63.Append(nonVisualShapeDrawingProperties63);
            nonVisualShapeProperties63.Append(applicationNonVisualDrawingProperties76);
            var shapeProperties63 = new ShapeProperties();

            var textBody63 = new TextBody();
            var bodyProperties63 = new A.BodyProperties();
            var listStyle63 = new A.ListStyle();

            var paragraph99 = new A.Paragraph();
            var endParagraphRunProperties58 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph99.Append(endParagraphRunProperties58);

            textBody63.Append(bodyProperties63);
            textBody63.Append(listStyle63);
            textBody63.Append(paragraph99);

            shape63.Append(nonVisualShapeProperties63);
            shape63.Append(shapeProperties63);
            shape63.Append(textBody63);

            var shape64 = new Shape();

            var nonVisualShapeProperties64 = new NonVisualShapeProperties();
            var nonVisualDrawingProperties77 = new NonVisualDrawingProperties() {
                Id = (UInt32Value) 7U,
                Name = "Slide Number Placeholder 6"
            };

            var nonVisualShapeDrawingProperties64 = new NonVisualShapeDrawingProperties();
            var shapeLocks64 = new A.ShapeLocks() {NoGrouping = true};

            nonVisualShapeDrawingProperties64.Append(shapeLocks64);

            var applicationNonVisualDrawingProperties77 = new ApplicationNonVisualDrawingProperties();
            var placeholderShape64 = new PlaceholderShape() {
                Type = PlaceholderValues.SlideNumber,
                Size = PlaceholderSizeValues.Quarter,
                Index = (UInt32Value) 12U
            };

            applicationNonVisualDrawingProperties77.Append(placeholderShape64);

            nonVisualShapeProperties64.Append(nonVisualDrawingProperties77);
            nonVisualShapeProperties64.Append(nonVisualShapeDrawingProperties64);
            nonVisualShapeProperties64.Append(applicationNonVisualDrawingProperties77);
            var shapeProperties64 = new ShapeProperties();

            var textBody64 = new TextBody();
            var bodyProperties64 = new A.BodyProperties();
            var listStyle64 = new A.ListStyle();

            var paragraph100 = new A.Paragraph();

            var field24 = new A.Field() {Id = "{D1985BBB-3714-4270-AB6D-2284F76009B2}", Type = "slidenum"};

            var runProperties86 = new A.RunProperties() {Language = "en-GB"};
            runProperties86.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            var text86 = new A.Text();
            text86.Text = "‹#›";

            field24.Append(runProperties86);
            field24.Append(text86);
            var endParagraphRunProperties59 = new A.EndParagraphRunProperties() {Language = "en-GB"};

            paragraph100.Append(field24);
            paragraph100.Append(endParagraphRunProperties59);

            textBody64.Append(bodyProperties64);
            textBody64.Append(listStyle64);
            textBody64.Append(paragraph100);

            shape64.Append(nonVisualShapeProperties64);
            shape64.Append(shapeProperties64);
            shape64.Append(textBody64);

            shapeTree13.Append(nonVisualGroupShapeProperties13);
            shapeTree13.Append(groupShapeProperties13);
            shapeTree13.Append(shape59);
            shapeTree13.Append(shape60);
            shapeTree13.Append(shape61);
            shapeTree13.Append(shape62);
            shapeTree13.Append(shape63);
            shapeTree13.Append(shape64);

            var commonSlideDataExtensionList13 = new CommonSlideDataExtensionList();

            var commonSlideDataExtension13 = new CommonSlideDataExtension() {
                Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
            };

            var creationId13 = new P14.CreationId() {Val = (UInt32Value) 134223542U};
            creationId13.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension13.Append(creationId13);

            commonSlideDataExtensionList13.Append(commonSlideDataExtension13);

            commonSlideData13.Append(shapeTree13);
            commonSlideData13.Append(commonSlideDataExtensionList13);

            var colorMapOverride12 = new ColorMapOverride();
            var masterColorMapping12 = new A.MasterColorMapping();

            colorMapOverride12.Append(masterColorMapping12);

            slideLayout11.Append(commonSlideData13);
            slideLayout11.Append(colorMapOverride12);

            slideLayoutPart11.SlideLayout = slideLayout11;
        }

        // Generates content of tableStylesPart1.
        private static void GenerateTableStylesPart1Content(TableStylesPart tableStylesPart1) {
            var tableStyleList1 = new A.TableStyleList() {Default = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"};
            tableStyleList1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            tableStylesPart1.TableStyleList = tableStyleList1;
        }

        // Generates content of viewPropertiesPart1.
        private static void GenerateViewPropertiesPart1Content(ViewPropertiesPart viewPropertiesPart1)
        {
            var viewProperties1 = new ViewProperties() {LastView = ViewValues.SlideThumbnailView};
            viewProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            viewProperties1.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            viewProperties1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            var normalViewProperties1 = new NormalViewProperties();
            var restoredLeft1 = new RestoredLeft() {Size = 15620};
            var restoredTop1 = new RestoredTop() {Size = 94660};

            normalViewProperties1.Append(restoredLeft1);
            normalViewProperties1.Append(restoredTop1);

            var slideViewProperties1 = new SlideViewProperties();

            var commonSlideViewProperties1 = new CommonSlideViewProperties();

            var commonViewProperties1 = new CommonViewProperties() {VariableScale = true};

            var scaleFactor1 = new ScaleFactor();
            var scaleX1 = new A.ScaleX() {Numerator = 116, Denominator = 100};
            var scaleY1 = new A.ScaleY() {Numerator = 116, Denominator = 100};

            scaleFactor1.Append(scaleX1);
            scaleFactor1.Append(scaleY1);
            var origin1 = new Origin() {X = -1578L, Y = -96L};

            commonViewProperties1.Append(scaleFactor1);
            commonViewProperties1.Append(origin1);

            var guideList1 = new GuideList();
            var guide1 = new Guide() {Orientation = DirectionValues.Horizontal, Position = 2160};
            var guide2 = new Guide() {Position = 2880};

            guideList1.Append(guide1);
            guideList1.Append(guide2);

            commonSlideViewProperties1.Append(commonViewProperties1);
            commonSlideViewProperties1.Append(guideList1);

            slideViewProperties1.Append(commonSlideViewProperties1);

            var notesTextViewProperties1 = new NotesTextViewProperties();

            var commonViewProperties2 = new CommonViewProperties();

            var scaleFactor2 = new ScaleFactor();
            var scaleX2 = new A.ScaleX() {Numerator = 1, Denominator = 1};
            var scaleY2 = new A.ScaleY() {Numerator = 1, Denominator = 1};

            scaleFactor2.Append(scaleX2);
            scaleFactor2.Append(scaleY2);
            var origin2 = new Origin() {X = 0L, Y = 0L};

            commonViewProperties2.Append(scaleFactor2);
            commonViewProperties2.Append(origin2);

            notesTextViewProperties1.Append(commonViewProperties2);
            var gridSpacing1 = new GridSpacing() {Cx = 72008L, Cy = 72008L};

            viewProperties1.Append(normalViewProperties1);
            viewProperties1.Append(slideViewProperties1);
            viewProperties1.Append(notesTextViewProperties1);
            viewProperties1.Append(gridSpacing1);

            viewPropertiesPart1.ViewProperties = viewProperties1;
        }

        // Generates content of extendedFilePropertiesPart1.
        private static void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1) {
            var properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt",
                "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            var totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "0";
            var words1 = new Ap.Words();
            words1.Text = "0";
            var application1 = new Ap.Application();
            application1.Text = "Microsoft Office PowerPoint";
            var presentationFormat1 = new Ap.PresentationFormat();
            presentationFormat1.Text = "On-screen Show (4:3)";
            var paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "0";
            var slides1 = new Ap.Slides();
            slides1.Text = "1";
            var notes1 = new Ap.Notes();
            notes1.Text = "0";
            var hiddenSlides1 = new Ap.HiddenSlides();
            hiddenSlides1.Text = "0";
            var multimediaClips1 = new Ap.MultimediaClips();
            multimediaClips1.Text = "0";
            var scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            var headingPairs1 = new Ap.HeadingPairs();

            var vTVector1 = new Vt.VTVector() {BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value) 4U};

            var variant1 = new Vt.Variant();
            var vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Theme";

            variant1.Append(vTLPSTR1);

            var variant2 = new Vt.Variant();
            var vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            var variant3 = new Vt.Variant();
            var vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Slide Titles";

            variant3.Append(vTLPSTR2);

            var variant4 = new Vt.Variant();
            var vTInt322 = new Vt.VTInt32();
            vTInt322.Text = "1";

            variant4.Append(vTInt322);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);
            vTVector1.Append(variant3);
            vTVector1.Append(variant4);

            headingPairs1.Append(vTVector1);

            var titlesOfParts1 = new Ap.TitlesOfParts();

            var vTVector2 = new Vt.VTVector() {BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value) 2U};
            var vTLPSTR3 = new Vt.VTLPSTR();
            vTLPSTR3.Text = "Office Theme";
            var vTLPSTR4 = new Vt.VTLPSTR();
            vTLPSTR4.Text = "PowerPoint Presentation";

            vTVector2.Append(vTLPSTR3);
            vTVector2.Append(vTLPSTR4);

            titlesOfParts1.Append(vTVector2);
            var company1 = new Ap.Company();
            company1.Text = "TNO";
            var linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            var sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            var hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            var applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "14.0000";

            properties1.Append(totalTime1);
            properties1.Append(words1);
            properties1.Append(application1);
            properties1.Append(presentationFormat1);
            properties1.Append(paragraphs1);
            properties1.Append(slides1);
            properties1.Append(notes1);
            properties1.Append(hiddenSlides1);
            properties1.Append(multimediaClips1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        private static void SetPackageProperties(OpenXmlPackage document) {
            document.PackageProperties.Creator = "Erik Vullings";
            document.PackageProperties.Title = "PowerPoint Presentation";
            document.PackageProperties.Revision = "2";
            document.PackageProperties.Created = XmlConvert.ToDateTime("2014-06-04T10:40:19Z",
                XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = XmlConvert.ToDateTime("2014-06-04T11:13:05Z",
                XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Erik Vullings";
        }

        #region Binary Data

        private const string ThumbnailPart1Data = "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCADAAQADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD+/iiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/ZbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbPwjZGz8I2Rs/CNkbEw6kAkAAHcIAAAAAAAAAADtBwAAYBmKCHMe/DB9GZQIlh4zMZkZngi4HmoxthmnCNoeojHTGbEI/R7ZMfAZuwgfHxAyDBrECEIfRzIpGs4IZB9/MhMAAAAAGs0DAAAAAAAAAAAAAOsIyx8kM5wa9QjtH1szuRr/CBAgkzPVGggJMiDKM/IaEglUIAE0DxscCXcgODQsGyUJmSBwNEgbLwm7IKc0ZRs5CQAAAAAAAAAAACEVNZ8bTAkiIU01uxtWCUUhhDXYG18JZyG7NfUbaQmJIfI1EhxzCawhKjYuHHwJziFhNkschgnwIZg2aByQCehFZGwAOJAJAAAAAAAAowk=";

        private static Stream GetBinaryDataStream(string base64String) {
            return new MemoryStream(Convert.FromBase64String(base64String));
        }

        #endregion
    }
}