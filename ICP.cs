using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Xml;

namespace ImageConversionPPT
{
    internal class ICP
    {
        public static string ImagePath { get; set; }

        private static int ImageCount { get; set; }


        /// <summary>
        /// Main
        /// </summary>
        public static async Task<bool> MainProcess()
        {
            await Task.Run(() =>
            {
                ExtractBasePPTMin();
                RemoveImages();
                AppXml();
                ContentTypesXml();
                PresentationXmlRels();
                SlideXmlRels();
                SlideXml();
                PresentationXml();
                ConversionPPT();
            });
            return true;
        }


        /// <summary>
        /// 解压BasePPTMin.zip
        /// </summary>
        private static void ExtractBasePPTMin()
        {
            System.IO.Compression.ZipFile.ExtractToDirectory(@"Resources\BasePPT.zip", ImagePath);
        }


        /// <summary>
        /// 生成.pptx
        /// </summary>
        private static void ConversionPPT()
        {
            string suffix = "";
            int suffixNum = 1;

            while (File.Exists(ImagePath + suffix + ".pptx")) {
                suffixNum += 1;
                suffix = "(" + suffixNum + ")";
            }

            System.IO.Compression.ZipFile.CreateFromDirectory(ImagePath, ImagePath + suffix + ".pptx");
            DirectoryInfo directory = new(ImagePath);
            directory.Delete(true);
        }


        /// <summary>
        /// 移动图片集
        /// </summary>
        private static void RemoveImages()
        {
            string destFolder = ImagePath + @"\ppt\media";

            if (Directory.Exists(destFolder) == false)
            {
                MessageBox.Show("未找到指定目录，执行失败。", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            DirectoryInfo directoryInfo = new(ImagePath);
            FileInfo[] files = directoryInfo.GetFiles();
            ImageCount = files.Length - 1;

            foreach (FileInfo file in files) // Directory.GetFiles(srcFolder)
            {
                if (file.Extension == ".jpg")
                {
                    file.MoveTo(Path.Combine(destFolder, file.Name), true);
                }
            }
        }


        /// <summary>
        /// ...\docProps\app.xml
        /// </summary>
        private static void AppXml()
        {
            XmlDocument xmlDoc = new();
            XmlNamespaceManager nsMgr = new(xmlDoc.NameTable);
            nsMgr.AddNamespace("ns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
            nsMgr.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

            xmlDoc.Load(ImagePath + @"\docProps\app.xml");

            XmlNode xmlRoot = xmlDoc.DocumentElement;
            xmlRoot.SelectSingleNode("//ns:Slides", nsMgr).InnerText = ImageCount.ToString();
            xmlRoot.SelectSingleNode("/ns:Properties/ns:HeadingPairs/vt:vector/vt:variant[last()]/vt:i4", nsMgr).InnerText = ImageCount.ToString();
            var modNode = (XmlElement)xmlRoot.SelectSingleNode("/ns:Properties/ns:TitlesOfParts/vt:vector", nsMgr);
            modNode.SetAttribute("size", (ImageCount + 4).ToString());
            for (int i = 1; i < ImageCount; i++)
            {
                modNode.AppendChild((XmlElement)modNode.LastChild.CloneNode(true));
            }

            xmlDoc.Save(ImagePath + @"\docProps\app.xml");
        }


        /// <summary>
        /// ...\[Content_Types].xml
        /// </summary>
        private static void ContentTypesXml()
        {
            XmlDocument xmlDoc = new();
            XmlNamespaceManager nsMgr = new(xmlDoc.NameTable);
            nsMgr.AddNamespace("ns", "http://schemas.openxmlformats.org/package/2006/content-types");

            xmlDoc.Load(ImagePath + @"\[Content_Types].xml");

            XmlNode xmlRoot = xmlDoc.DocumentElement;
            for (int i = 2; i <= ImageCount; i++)
            {
                XmlElement newElem = xmlDoc.CreateElement("Override", xmlDoc.DocumentElement.NamespaceURI);
                newElem.SetAttribute("PartName", "/ppt/slides/slide" + i + ".xml");
                newElem.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");
                xmlRoot.AppendChild(newElem);
            }

            xmlDoc.Save(ImagePath + @"\[Content_Types].xml");
        }


        /// <summary>
        /// ...\ppt\_rels\presentation.xml.rels
        /// </summary>
        private static void PresentationXmlRels()
        {
            XmlDocument xmlDoc = new();
            XmlNamespaceManager nsMgr = new(xmlDoc.NameTable);
            nsMgr.AddNamespace("ns", "http://schemas.openxmlformats.org/package/2006/relationships");

            xmlDoc.Load(ImagePath + @"\ppt\_rels\presentation.xml.rels");

            XmlNode xmlRoot = xmlDoc.DocumentElement;
            ((XmlElement)xmlRoot.SelectSingleNode("/ns:Relationships/ns:Relationship[@Target = 'presProps.xml']", nsMgr)).SetAttribute("Id", "rId" + (ImageCount + 2));
            ((XmlElement)xmlRoot.SelectSingleNode("/ns:Relationships/ns:Relationship[@Target = 'viewProps.xml']", nsMgr)).SetAttribute("Id", "rId" + (ImageCount + 3));
            ((XmlElement)xmlRoot.SelectSingleNode("/ns:Relationships/ns:Relationship[@Target = 'theme/theme1.xml']", nsMgr)).SetAttribute("Id", "rId" + (ImageCount + 4));
            ((XmlElement)xmlRoot.SelectSingleNode("/ns:Relationships/ns:Relationship[@Target = 'tableStyles.xml']", nsMgr)).SetAttribute("Id", "rId" + (ImageCount + 5));

            for (int i = 2; i <= ImageCount; i++)
            {
                XmlElement newElem = xmlDoc.CreateElement("Relationship", xmlDoc.DocumentElement.NamespaceURI);
                newElem.SetAttribute("Id", "rId" + (i + 1));
                newElem.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide");
                newElem.SetAttribute("Target", "slides/slide" + i + ".xml");
                xmlRoot.AppendChild(newElem);
            }

            xmlDoc.Save(ImagePath + @"\ppt\_rels\presentation.xml.rels");
        }


        /// <summary>
        /// ...\ppt\slides\_rels\slide.xml.rels
        /// </summary>
        private static void SlideXmlRels()
        {
            for (int i = 2; i <= ImageCount; i++)
            {
                XmlDocument xmlDoc = new();
                xmlDoc.AppendChild(xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", "yes"));

                XmlElement relationships = xmlDoc.CreateElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                xmlDoc.AppendChild(relationships);

                XmlElement relationship2 = xmlDoc.CreateElement("Relationship", xmlDoc.DocumentElement.NamespaceURI);
                relationship2.SetAttribute("Id", "rId" + (i + 1).ToString());
                relationship2.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                relationship2.SetAttribute("Target", "../media/image" + i + ".jpg");
                relationships.AppendChild(relationship2);

                XmlElement relationship1 = xmlDoc.CreateElement("Relationship", xmlDoc.DocumentElement.NamespaceURI);
                relationship1.SetAttribute("Id", "rId1");
                relationship1.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout");
                relationship1.SetAttribute("Target", "../slideLayouts/slideLayout7.xml");
                relationships.AppendChild(relationship1);

                xmlDoc.Save(ImagePath + @"\ppt\slides\_rels\slide" + i + ".xml.rels");
            }
        }


        /// <summary>
        /// ...\ppt\slides\slide.xml
        /// </summary>
        private static void SlideXml()
        {
            XmlDocument xmlDoc = new();
            XmlNamespaceManager nsMgr = new(xmlDoc.NameTable);
            nsMgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            nsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            nsMgr.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            nsMgr.AddNamespace("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            xmlDoc.Load(ImagePath + @"\ppt\slides\slide1.xml");

            XmlNode xmlRoot = xmlDoc.DocumentElement;
            for (int i = 2; i <= ImageCount; i++)
            {
                ((XmlElement)xmlRoot.SelectSingleNode("//a:blip", nsMgr)).SetAttribute("r:embed", "rId" + (i + 1).ToString());
                Random rnd = new();
                ((XmlElement)xmlRoot.SelectSingleNode("//p14:creationId", nsMgr)).SetAttribute("val", "202108" + rnd.Next(1000, 10000));

                xmlDoc.Save(ImagePath + @"\ppt\slides\slide" + i + ".xml");
            }
        }


        /// <summary>
        /// ...\ppt\presentation.xml
        /// </summary>
        private static void PresentationXml()
        {
            XmlDocument xmlDoc = new();
            XmlNamespaceManager nsMgr = new(xmlDoc.NameTable);
            nsMgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            nsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            nsMgr.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            xmlDoc.Load(ImagePath + @"\ppt\presentation.xml");

            XmlNode xmlRoot = xmlDoc.DocumentElement;
            XmlNode modNode = xmlRoot.SelectSingleNode("//p:sldIdLst", nsMgr);
            for (int i = 3; i <= ImageCount + 1; i++)
            {
                XmlElement newElem = xmlDoc.CreateElement("p:sldId", xmlDoc.DocumentElement.NamespaceURI);
                newElem.SetAttribute("id", (i + 254).ToString());
                XmlAttribute newAttribute = xmlDoc.CreateAttribute("r", "id", nsMgr.LookupNamespace("r"));
                newAttribute.Value = "rId" + i;
                newElem.Attributes.Append(newAttribute);
                modNode.AppendChild(newElem);
            }

            xmlDoc.Save(ImagePath + @"\ppt\presentation.xml");
        }


    }
}
