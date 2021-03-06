﻿using System;
using System.IO;
using NPOI.OpenXml4Net.OPC;
using Xunit;

namespace npoi.test.ooxml
{
    public class CreateBasicOOXMLFile
    {
        [Fact]
        public void Main()
        {
            //create ooxml file in memory
            OPCPackage p = OPCPackage.Create(new MemoryStream());

            //create package parts
            PackagePartName pn1 = new PackagePartName(new Uri("/a/abcd/e", UriKind.Relative), true);
            //if (!p.ContainPart(pn1))
                //p.CreatePart(pn1, MediaTypeNames.Text.Plain);

            PackagePartName pn2 = new PackagePartName(new Uri("/b/test.xml", UriKind.Relative), true);
            //if (!p.ContainPart(pn2))
                //p.CreatePart(pn2, MediaTypeNames.Text.Xml);

            //save file 
            p.Save("test.zip");

            //don't forget to close it
            p.Close();
        }
    }
}
