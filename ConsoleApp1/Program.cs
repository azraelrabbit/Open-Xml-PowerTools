using System;
using System.Drawing.Imaging;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            var docpath = "d:\\25义乌机场风险分析.docx";

            var toPath = "d:\\25义乌机场风险分析.html";

            var fi=new FileInfo(docpath);


            var imgsPath = Path.Combine(fi.DirectoryName, fi.Name + "_imgs");

            if (!Directory.Exists(imgsPath))
            {
                Directory.CreateDirectory(imgsPath);
            }


            int imgCounter = 0;

            var allbytes = File.ReadAllBytes(docpath);

            using (var fs = new MemoryStream())
            {
                fs.Write(allbytes,0,allbytes.Length);

                HtmlConverterSettings sett=new HtmlConverterSettings()
                {
                    PageTitle = fi.Name,
                    ImageHandler=new Func<ImageInfo, XElement>(img =>
                    {
                        imgCounter++;
                        ImageFormat imgfmt=null;
                        var extension = img.ContentType.Split('/')[1].ToLower();

                        switch (extension)
                        {
                                 case "png":
                                     imgfmt=ImageFormat.Png;
                                     break;
                            case "jpeg":
                                imgfmt=ImageFormat.Jpeg;
                                break;
                            case "bmp":
                                imgfmt=ImageFormat.Bmp;
                                break;
                            case "tiff":
                                imgfmt=ImageFormat.Tiff;
                                break;
                        }

                        if (imgfmt == null)
                        {
                            return null;
                        }

                        var imgPath = Path.Combine(imgsPath, img.AltText + imgCounter.ToString() + "." + extension);
                        img.Bitmap.Save(imgPath,imgfmt);
                        var imgel=new XElement(Xhtml.img,new XAttribute(NoNamespace.src,imgPath),img.ImgStyleAttribute,img.AltText!=null?new XAttribute(NoNamespace.alt,img.AltText) :null );
                        return imgel;
                    } )
                };
                using (var doc = WordprocessingDocument.Open(fs, true))
                {
                    var htmlConverter = HtmlConverter.ConvertToHtml(doc, sett);

                    File.WriteAllText(toPath,htmlConverter.ToStringNewLineOnAttributes());
                }
            }


            Console.WriteLine("Hello World!");

            Console.ReadLine();
        }
    }
}
