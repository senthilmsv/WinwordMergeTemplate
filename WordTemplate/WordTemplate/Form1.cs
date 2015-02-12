using Novacode;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordTemplate
{
    /*
     * http://www.codeproject.com/Articles/660478/Csharp-Create-and-Manipulate-Word-Documents-Progra
     * */
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            CreateRejectionLetter("%APPLICANT%", "Senthil");
        }

        public void CreateRejectionLetter(string applicantField, string applicantName)
        {
            // We will need a file name for our output file (change to suit your machine):
            string fileNameTemplate = @"C:\temp\Rejection-Letter-{0}-{1}.docx";

            // Let's save the file with a meaningful name, including the applicant name and the letter date:
            string outputFileName = string.Format(fileNameTemplate, applicantName, DateTime.Now.ToString("MM-dd-yy"));

            // Grab a reference to our document template:
            DocX letter = this.GetRejectionLetterTemplate();

            // Perform the replace:
            letter.ReplaceText(applicantField, applicantName);

            letter.ReplaceText("[Field1]", "First Replacement");

            letter.ReplaceText("[Field2]", "Second Replacement");

            // Save as New filename:
            letter.SaveAs(outputFileName);

            // Open in word:
            Process.Start("WINWORD.EXE", "\"" + outputFileName + "\"");
        }

        private DocX GetRejectionLetterTemplate()
        {
            // Adjust the path so suit your machine:
            string fileName = @"D:\Users\John\Documents\DocXExample.docx";

            // Set up our paragraph contents:
            string headerText = "Find and Replace Text Using DocX - Merge Templating, Anyone?";
            string letterBodyText = DateTime.Now.ToShortDateString();

            string paraOne = ""
                + "I have a word document which contains only one page filled with text and graphic. "
                + "The page also contains some placeholders like [Field1],[Field2],..., etc. "
                + "I get data from database and I want to open this document and fill placeholders with some data. "
            + "For each data row I want to open this document, fill placeholders with row's data and then concatenate all created documents into one document. "
            + "What is the best and simpliest way to do this?";

            string paraTwo = ""
                + "Dear %APPLICANT%" + Environment.NewLine + Environment.NewLine
                + "I am writing to thank you for your resume. Unfortunately, your skills and "
                + "experience do not match our needs at the present time. We will keep your "
                + "resume in our circular file for future reference. Don't call us, we'll call you. "
                + Environment.NewLine + Environment.NewLine
                + "Sincerely, "
                + Environment.NewLine + Environment.NewLine
                + "Jim Smith, Corporate Hiring Manager";

            // Title Formatting:
            var titleFormat = new Formatting();
            titleFormat.FontFamily = new System.Drawing.FontFamily("Arial Black");
            titleFormat.Size = 18D;
            titleFormat.Position = 12;

            // Body Formatting
            var paraFormat = new Formatting();
            paraFormat.FontFamily = new System.Drawing.FontFamily("Calibri");
            paraFormat.Size = 10D;
            titleFormat.Position = 12;

            // Create the document in memory:
            var doc = DocX.Create(fileName);

            // Insert each prargraph, with appropriate spacing and alignment:
            Paragraph title = doc.InsertParagraph(headerText, false, titleFormat);
            title.Alignment = Alignment.center;

            doc.InsertParagraph(Environment.NewLine);
            Paragraph letterBody = doc.InsertParagraph(letterBodyText, false, paraFormat);
            letterBody.Alignment = Alignment.both;

            doc.InsertParagraph(Environment.NewLine);
            Paragraph letterBody1 = doc.InsertParagraph(paraOne, false, paraFormat);
            letterBody1.Alignment = Alignment.both;

            doc.InsertParagraph(Environment.NewLine);
            doc.InsertParagraph(paraTwo, false, paraFormat);

            return doc;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            CreateSampleDocument();
        }

        public void CreateSampleDocument()
        {
            // Modify to siut your machine:
            string fileName =  @"c:\temp\DocXExample.docx";

            //// Create a document in memory:
            //var doc = DocX.Create(fileName);

            //// Insert a paragrpah:
            //doc.InsertParagraph("This is my first paragraph");

            //// Save to the output directory:
            //doc.Save();

            //// Open in Word:
            //Process.Start("WINWORD.EXE", fileName);

            string headlineText = "C#: Create and Manipulate Word Documents Programmatically Using DocX";
            string paraOne = ""
                + "I have a word document which contains only one page filled with text and graphic. "
                + "The page also contains some placeholders like [Field1],[Field2],..., etc. "
                + "I get data from database and I want to open this document and fill placeholders with some data. "
            + "For each data row I want to open this document, fill placeholders with row's data and then concatenate all created documents into one document. "
            + "What is the best and simpliest way to do this?";

            // A formatting object for our headline:
            var headLineFormat = new Formatting();
            headLineFormat.FontFamily = new System.Drawing.FontFamily("Arial Black");
            headLineFormat.Size = 18D;
            headLineFormat.Position = 12;

            // A formatting object for our normal paragraph text:
            var paraFormat = new Formatting();
            paraFormat.FontFamily = new System.Drawing.FontFamily("Calibri");
            paraFormat.Size = 10D;

            // Create the document in memory:
            var doc = DocX.Create(fileName);

            // Insert the now text obejcts;
            doc.InsertParagraph(headlineText, false, headLineFormat);
            doc.InsertParagraph(paraOne, false, paraFormat);

            // Save to the output directory:
            doc.Save();

            // Open in Word:
            Process.Start("WINWORD.EXE", fileName);
        }
    }
}
