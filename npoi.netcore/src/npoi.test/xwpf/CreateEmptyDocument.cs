using System.IO;
using System.Text;
using NPOI.XWPF.UserModel;
using Xunit;

namespace npoi.test.xwpf
{
    public class CreateEmptyDocument
    {
        [Fact]
        public void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            XWPFDocument doc = new XWPFDocument();
            doc.CreateParagraph();
            using (FileStream sw = File.Create("blank.docx"))
            {
                doc.Write(sw);
            }
        }
    }
}
