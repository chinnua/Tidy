using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using TidyManaged;
using word = Microsoft.Office.Interop.Word;

namespace Tidy
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                word.Application wrdApp = new word.Application();
                wrdApp.Visible = false;
                wrdApp.DisplayAlerts = word.WdAlertLevel.wdAlertsNone;
                word.Document wrdDoc = wrdApp.Documents.Open(@"C:\Users\Chinnu\Desktop\Tidy\Input.docx", ReadOnly: true);
                wrdDoc.SaveAs2(@"C:\Users\Chinnu\Desktop\Tidy\Input.html", FileFormat: word.WdSaveFormat.wdFormatFilteredHTML);
                wrdDoc.Close(SaveChanges: false);
                wrdApp.Quit(SaveChanges: false);

                string html = File.ReadAllText(@"C:\Users\Chinnu\Desktop\Tidy\Input.html");
                var premailerResult = PreMailer.Net.PreMailer.MoveCssInline(html);
                string inlinedHTML = premailerResult.Html;
                using (Document doc = Document.FromString(inlinedHTML))
                {
                    doc.ShowWarnings = false;
                    doc.Quiet = true;
                    doc.OutputXhtml = true;
                    doc.AddTidyMetaElement = false;
                    doc.DocType = DocTypeMode.Strict;
                    doc.MakeBare = true;
                    //doc.MakeClean = true; ==> This converts inline CSS to Style Tags
                    doc.CleanAndRepair();
                    doc.Save(@"C:\Users\Chinnu\Desktop\Tidy\Output.htm");                    
                }
                MessageBox.Show("Done");

            }
            catch (Exception ex)
            {

                throw;
            }

        }
    }
}
