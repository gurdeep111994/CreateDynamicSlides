//using GemBox.Presentation;
using System;

using Syncfusion.Presentation;
using System.IO;
using System.Collections.Generic;

namespace CreatePPT
{
    class Program
    {
        static void Main(string[] args)
        {
            List<PageType> pageTypes = new List<PageType>()
            {
                new PageType()
                {
                    Heading="React JS",
                    Type="master",
                    Paragraphs = new List<Paragraph>()
                    {
                          new Paragraph(){ Text = "Hello and Welcome"},
                          new Paragraph(){ Text = "To"},
                          new Paragraph(){ Text = "{ 360 coding syntaxes }"},
                    }
                },
                new PageType()
                {
                    Heading="What is React?",
                    Type="child",
                    Paragraphs = new List<Paragraph>()
                    {
                          new Paragraph(){ Text = "It is JavaScript open source library for building( front-end applications )user interfaces."},
                          new Paragraph(){ Text = "Don't confuse it with a framework"},
                          new Paragraph(){ Text = "Its main focus is on building UI"},
                          new Paragraph(){ Text = "It does not concern with other part of application i.e. routing, http requests"},
                    }
                },
                new PageType()
                {
                    Heading="Why we need to learn React?",
                    Type="child",
                    Paragraphs = new List<Paragraph>()
                    {
                          new Paragraph(){ Text = "Created and maintained by Facebook"},
                          new Paragraph(){ Text = "More than 100k stars on GitHub "},
                          new Paragraph(){ Text = "Huge community"},
                          new Paragraph(){ Text = "In demand skillset"},
                    }
                },
                new PageType()
                {
                    Heading="Why it is a better choice?",
                    Type="child",
                    Paragraphs = new List<Paragraph>()
                    {
                          new Paragraph(){ Text = "Component Based Architecture"},
                          new Paragraph(){ Text = "Reusable code"},
                          new Paragraph(){ Text = "Efficient and fast"},
                          new Paragraph(){ Text = "Works in browser"},
                          new Paragraph(){ Text = "Enterprise application ability to use code "},
                    }
                },
                new PageType()
                {
                    Heading="More on React",
                    Type="child",
                    Paragraphs = new List<Paragraph>()
                    {
                          new Paragraph(){ Text = "react will handle efficiently updating and rendering of the components"},
                          new Paragraph(){ Text = "DOM updates are handled gracefully in React"},
                          new Paragraph(){ Text = "seamlessly integrate react into any of your applications"},
                          new Paragraph(){ Text = "portion of your page , complete page or even an entire application itself."},
                    }
                },
            };

           CreatePPT(pageTypes);
           // CreatePPT2();
        }

        private static void CreatePPT(List<PageType> pageTypes)
        {
            try
            {
                IPresentation powerpointDoc = Presentation.Create();

                foreach (var page in pageTypes)
                {
                    if (page.Type.ToLower() == "master" || page.Type.ToLower() == "child")
                    {
                        ISlide slide = powerpointDoc.Slides.Add(Syncfusion.Presentation.SlideLayoutType.Blank);
                        slide.Background.Fill.FillType = FillType.Solid;
                        slide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(184, 27, 232);
                        for (int i = 0; i < 10; i++)
                        {
                            IShape titleShape = slide.AddTextBox(53.22, 70.73, 874.19, 77.70);
                            titleShape.TextBody.AddParagraph(page.Heading).HorizontalAlignment = HorizontalAlignmentType.Center;
                        }
                        double top = 140.73;
                        if (page.Type.ToLower() == "master")
                        {
                            
                            foreach (var paragraph in page.Paragraphs)
                            {
                                IShape paraText = slide.AddTextBox(53.22, top, 874.19, 77.70);
                                paraText.TextBody.AddParagraph(paragraph.Text).HorizontalAlignment = HorizontalAlignmentType.Center;
                                top = top + 40.00;
                            }
                        } else if(page.Type.ToLower() == "child")
                        {
                            int paragrapghCount = page.Paragraphs.Count;
                            if(paragrapghCount > 0)
                            {
                                int height = 56 * paragrapghCount;
                                IShape bulletPointsShape = slide.AddTextBox(53.22, top, 874.19, height);
                                foreach (var paragraph in page.Paragraphs)
                                {
                                    IParagraph bulletPAra = bulletPointsShape.TextBody.AddParagraph(paragraph.Text);
                                    bulletPAra.ListFormat.Type = Syncfusion.Presentation.ListType.Bulleted;
                                    bulletPAra.LeftIndent = 35;
                                    bulletPAra.FirstLineIndent = -35;
                                }
                            }
                        }

                        IShape stampShape = slide.Shapes.AddShape(AutoShapeType.Explosion1, 48.93, 430.71, 104.13, 80.54);
                        stampShape.Fill.FillType = FillType.None;
                        stampShape.TextBody.AddParagraph("360").HorizontalAlignment = HorizontalAlignmentType.Center;
                    }
                }
                var customDate = DateTime.Now;
                FileStream outputStream = new FileStream($@"Sample_{customDate.Day}_{customDate.Month}_{customDate.Year}_{customDate.Hour}_{customDate.Minute}_{customDate.Second}_{customDate.Millisecond}.pptx", FileMode.Create); powerpointDoc.Save(outputStream);
                outputStream.Dispose();
                powerpointDoc.Close();
            }
            catch (Exception  ex)
            {
                throw;
            }
        }

    }
}