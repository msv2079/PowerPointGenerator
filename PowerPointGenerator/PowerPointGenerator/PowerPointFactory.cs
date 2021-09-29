using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointGenerator
{
    public class PowerPointFactory
    {
        private readonly string presentationFile;

        public PowerPointFactory(string fileName)
        {
            presentationFile = fileName;
            Template.CreatePackage(presentationFile);
        }

        /// <summary>
        ///     Create new slides, each with an image and a title.
        /// </summary>
        /// <param name="imagePaths"></param>
        /// <param name="titles"></param>
        public void CreateTitleAndImageSlides(List<string> imagePaths, List<string> titles = null)
        {
            if (imagePaths == null) throw new ArgumentNullException("imagePaths");
            if (titles != null && titles.Count != imagePaths.Count)
                throw new ArgumentOutOfRangeException("titles", "Titles must contain as many elements as images.");

            if (titles == null)
            {
                titles = imagePaths.Select(Path.GetFileNameWithoutExtension).ToList();
            }

            // Open the source document as read/write. 
            using (var presentationDocument = PresentationDocument.Open(presentationFile, true))
            {
                var position = 0;
                for (var i = 0; i < imagePaths.Count; i++)
                {
                    var path = imagePaths[i];
                    var title = titles[i];
                    var slide = InsertNewSlide(presentationDocument, position++, title);
                    InsertImageInSlide(slide, path, ImageExtension(path));
                }
                DeleteTemplateSlide(presentationDocument);
            }
        }

        /// <summary>
        ///     Determine the appropriate image extension to use.
        /// </summary>
        /// <param name="imagefilePath"></param>
        /// <returns></returns>
        private static string ImageExtension(string imagefilePath)
        {
            var imageExt = Path.GetExtension(imagefilePath);
            if (string.IsNullOrEmpty(imageExt))
                throw new ArgumentException("Image extension should be either jpg, jpeg or png.", "imagefilePath");
            return imageExt.Equals("jpg", StringComparison.OrdinalIgnoreCase) ||
                   imageExt.Equals("jpeg", StringComparison.OrdinalIgnoreCase)
                ? "image/jpeg"
                : "image/png";
        }

        /// <summary>
        ///     Insert the specified slide into the presentation at the specified position.
        /// </summary>
        /// <param name="presentationDocument"></param>
        /// <param name="position"></param>
        /// <param name="slideTitle"></param>
        /// <returns></returns>
        private static P.Slide InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
        {
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            if (slideTitle == null)
            {
                throw new ArgumentNullException("slideTitle");
            }

            var presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart == null)
            {
                throw new InvalidOperationException("The presentation document is empty.");
            }

            // Declare and instantiate a new slide.
            var slide = new P.Slide(new P.CommonSlideData(new P.ShapeTree()));
            uint drawingObjectId = 1;

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            var nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new P.NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new P.NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new P.GroupShapeProperties());

            // Declare and instantiate the title shape of the new slide.
            var titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());

            drawingObjectId++;

            // Specify the required shape properties for the title shape. 
            titleShape.NonVisualShapeProperties =
                new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties { Id = drawingObjectId, Name = "Title" },
                    new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape
                    {
                        Type = P.PlaceholderValues.Title
                    }));
            titleShape.ShapeProperties = new P.ShapeProperties();

            // Specify the text of the title shape.
            titleShape.TextBody = new P.TextBody(new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text { Text = slideTitle })));

            #region body shape

            // TODO If you need a slide with a body, make sure that the generated template contains a body, and uncomment the following lines. 
            //// Declare and instantiate the body shape of the new slide.
            //var bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
            //drawingObjectId++;

            //// Specify the required shape properties for the body shape.
            //bodyShape.NonVisualShapeProperties = new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties { Id = drawingObjectId, Name = "Content Placeholder" },
            //    new P.NonVisualShapeDrawingProperties(new ShapeLocks { NoGrouping = true }),
            //    new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Index = 1 }));
            //bodyShape.ShapeProperties = new P.ShapeProperties();

            //// Specify the text of the body shape.
            //bodyShape.TextBody = new P.TextBody(new BodyProperties(), new ListStyle(), new Paragraph());

            #endregion body shape

            // Create the slide part for the new slide.
            var slidePart = presentationPart.AddNewPart<SlidePart>();

            // Save the new slide part.
            slide.Save(slidePart);

            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            var slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            P.SlideId prevSlideId = null;

            foreach (var slideId in slideIdList.ChildElements.Cast<P.SlideId>())
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0)
                {
                    prevSlideId = slideId;
                }
            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((P.SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            // Insert the new slide into the slide list after the previous slide.
            var newSlideId = slideIdList.InsertAfter(new P.SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();

            return GetSlideByRelationShipId(presentationPart, newSlideId.RelationshipId);
        }

        /// <summary>
        ///     Get Slide By RelationShip ID
        /// </summary>
        /// <param name="presentationPart">Presentation Part</param>
        /// <param name="relationshipId">Relationship ID</param>
        /// <returns>Slide Object</returns>
        private static P.Slide GetSlideByRelationShipId(PresentationPart presentationPart, StringValue relationshipId)
        {
            // Get Slide object by Relationship ID
            var slidePart = presentationPart.GetPartById(relationshipId) as SlidePart;
            return slidePart != null ? slidePart.Slide : null;
        }

        /// <summary>
        ///     Insert Image into Slide
        /// </summary>
        /// <param name="slide">The slide</param>
        /// <param name="imagePath">Image Path</param>
        /// <param name="imageExt">Image Extension</param>
        private static void InsertImageInSlide(P.Slide slide, string imagePath, string imageExt)
        {
            // Creates an Picture instance and adds its children. 
            var picture = new P.Picture();
            var embedId = "rId" + (slide.Elements().Count() + 915);
            var nonVisualPictureProperties = new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = (UInt32Value)4U, Name = "Picture 5" },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties());

            var blipFill = new P.BlipFill();
            var blip = new A.Blip { Embed = embedId };

            // Creates an BlipExtensionList instance and adds its children 
            var blipExtensionList = new A.BlipExtensionList();
            var blipExtension     = new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            var useLocalDpi = new UseLocalDpi { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension.Append(useLocalDpi);
            blipExtensionList.Append(blipExtension);
            blip.Append(blipExtensionList);

            var stretch       = new A.Stretch();
            var fillRectangle = new A.FillRectangle();
            stretch.Append(fillRectangle);

            blipFill.Append(blip);
            blipFill.Append(stretch);

            // Generates content of imagePart. 
            //http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2009/09/06/7429.aspx
            var newSlidePart = slide.SlidePart;
            var imagePart = newSlidePart.AddNewPart<ImagePart>(imageExt, embedId);

            using (var imageStream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(imageStream);
            }
            
            //need to readjust the image container size and position depending on the image's aspect ratio
            var image = Image.FromFile(imagePath);
            var aspectRatio = (double)image.Width / image.Height;

            // Compute the image's offset on the page (in x and y), and its width cx and height cy.
            // Note that sizes are expressed in EMU (English Metric Units)
            // const int emusPerCm = 360000;
            var cy = 5029200L;
            var cx = (long)(cy * aspectRatio);
            if (cx > 8229600L)
            {
                cx = 8229600L;
                cy = (long)(cx / aspectRatio);
            }

            // Creates an ShapeProperties instance and adds its children. 
            var shapeProperties = new P.ShapeProperties();
            var transform2D     = new A.Transform2D();
            var offset          = new A.Offset { X = (9144000L - cx) / 2, Y = 1524000L };
            var extents         = new A.Extents { Cx = cx, Cy = cy };

            transform2D.Append(offset);
            transform2D.Append(extents);

            var presetGeometry  = new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle };
            var adjustValueList = new A.AdjustValueList();

            presetGeometry.Append(adjustValueList);

            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);

            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFill);
            picture.Append(shapeProperties);

            slide.CommonSlideData.ShapeTree.AppendChild(picture);
        }

        /// <summary>
        ///     Delete the template slide from the template presentation
        /// </summary>
        private static void DeleteTemplateSlide(PresentationDocument doc)
        {
            //delete the template slide and any references
            var slideIdList = doc.PresentationPart.Presentation.SlideIdList;

            foreach (var openXmlElement in slideIdList.ChildElements)
            {
                var slideId = (P.SlideId)openXmlElement;
                if (slideId.RelationshipId.Value.Equals("rId2")) slideIdList.RemoveChild(slideId);
            }

            var slideTemplate = doc.PresentationPart.SlideParts.First();
            doc.PresentationPart.DeletePart(slideTemplate);
            doc.PresentationPart.Presentation.Save();
        }
    }
}