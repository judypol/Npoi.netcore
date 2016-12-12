using System.IO;
using npoi.core.Corefx;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using Xunit;

namespace npoi.test.poifs
{
    public class CreateCustomProperties
    {
        [Fact]
        public void Main()
        {
            POIFSFileSystem fs = new POIFSFileSystem();

            //get the root directory
            DirectoryEntry dir = fs.Root;

            //create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            CustomProperties customProperties = dsi.CustomProperties;
            if (customProperties == null)
                customProperties = new CustomProperties();
            customProperties.Put("测试", "value A");
            customProperties.Put("BB", "value BB");
            customProperties.Put("CCC", "value CCC");
            dsi.CustomProperties = customProperties;
            //Write the stream data of the DocumentSummaryInformation entry to the root directory
            dsi.Write(dir, DocumentSummaryInformation.DEFAULT_STREAM_NAME);


            //create a POIFS file called Foo.poifs
            FileStream output = new FileStream("Foo.xls", FileMode.OpenOrCreate);
            fs.WriteFileSystem(output);
            output.Close();
        }
    }
}
