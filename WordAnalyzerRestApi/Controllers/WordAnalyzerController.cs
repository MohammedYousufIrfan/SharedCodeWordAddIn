using Aspose.Words;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Text;

namespace WordAnalyzerRestApi.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class WordAnalyzerController : ControllerBase
    {
        [HttpGet("unicode")]
        public ActionResult<string> GetUnicode(string value)
        {
            if (value == null)
            {
                return BadRequest();
            }
            return SharedCodeWordLibrary.WordOperations.GetUnicode(value);
        }

        [HttpGet("charcount")]
        public ActionResult<string> GetCharCount(string value)
        {
            if (value == null)
            {
                return BadRequest();
            }
            return SharedCodeWordLibrary.WordOperations.GetCharCount(value);
        }

        [HttpGet("wordcount")]
        public ActionResult<string> GetWordCount(string value)
        {
            if (value == null)
            {
                return BadRequest();
            }
            return SharedCodeWordLibrary.WordOperations.GetWordCount(value);
        }
        [HttpPost("addtoc")]
        public ActionResult<string> AddToc([FromBody] WordOOXML wordOOXML)
        {
            if (wordOOXML.XmlData == null)
            {
                return BadRequest();
            }
            // string dataDir = @"C:\Users\Yousuf.Irfan\Desktop\web (2).xml";

            MemoryStream mStrm = new MemoryStream(Encoding.UTF8.GetBytes(wordOOXML.XmlData));
            Document doc = new Document(mStrm);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");


            doc.UpdateFields();
            string filepath = Path.GetTempPath() + "output2.xml";
            doc.Save(filepath);
           var ouputxml=System.IO.File.ReadAllText(filepath);
            return ouputxml;
        }
    }
    public class WordOOXML
    {
        public string XmlData { get; set; }
    }
}
