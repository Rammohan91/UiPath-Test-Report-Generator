using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace TestReportGenerator
{
    public class EndTest : CodeActivity
    {

        // Arguments of End Test Activity
        private InArgument<String> statusValue = "PASSED";

        [Category("Status")]
        [RequiredArgument]
        public InArgument<String> Status
        {
            get { return this.statusValue; }
            set { this.statusValue = value; }
        }

        private static readonly Regex Regex = new Regex("[^a-zA-Z0-9 -]");


        protected override void Execute(CodeActivityContext context)
        {
            if (File.Exists("Temp.xml"))
            {

                XDocument doc = XDocument.Load("Temp.xml");

                XElement updateStartedElement = doc.Element("Test-Suite").Element("Result").Element("Status");
                updateStartedElement.Value = Status.Get(context);

                XElement addEndedElement = doc.Element("Test-Suite").Element("Result");
                addEndedElement.Add(new XElement("Ended", DateTime.Now.ToString()));

                string testScenario = doc.Element("Test-Suite").Element("Result").Element("TestScenario").Value;
                string testCase = doc.Element("Test-Suite").Element("Result").Element("TestName").Value;
                string startedTime = doc.Element("Test-Suite").Element("Result").Element("Started").Value;
                string endedTime = doc.Element("Test-Suite").Element("Result").Element("Ended").Value;
                string status = doc.Element("Test-Suite").Element("Result").Element("Status").Value;

                doc.Save("Temp.xml");

                // Create or Append to Testing-Report.XML
                if (!File.Exists("Testing-Report.xml"))
                {
                    XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
                    xmlWriterSettings.Indent = true;
                    xmlWriterSettings.NewLineOnAttributes = true;
                    using (XmlWriter xmlWriter = XmlWriter.Create("Testing-Report.xml", xmlWriterSettings))
                    {
                        xmlWriter.WriteStartDocument();
                        xmlWriter.WriteStartElement("Test-Suite");
                        xmlWriter.WriteStartElement("Result");
                        xmlWriter.WriteElementString("TestScenario", testScenario);
                        xmlWriter.WriteElementString("TestName", testCase);
                        xmlWriter.WriteElementString("Started", startedTime);
                        xmlWriter.WriteElementString("Ended", endedTime);
                        xmlWriter.WriteElementString("Status", status);
                        xmlWriter.WriteEndElement();

                        xmlWriter.WriteEndElement();
                        xmlWriter.WriteEndDocument();
                        xmlWriter.Flush();
                        xmlWriter.Close();

                        XDocument doc_Main = XDocument.Load("Testing-Report.xml");
                        XElement element = doc_Main.Element("Test-Suite");

                        // Add Atrributes
                        CheckAndUpdateAttributes(element);

                        //Increment Total Attributes
                        IncrementTotalANDPassedAttributes(element);

                        //Save Element
                        element.Save("Testing-Report.xml");
                    }
                }
                else
                {
                    XDocument doc_Main = XDocument.Load("Testing-Report.xml");
                    XElement element = doc_Main.Element("Test-Suite");

                    // Increment Passed & Failures
                    IncrementTotalANDPassedAttributes(element);

                    //Save element
                    element.Save("Testing-Report.xml");

                    //Add Element
                    element.Add(new XElement("Result", new XElement("TestScenario", testScenario),
                        new XElement("TestName", testCase),
                        new XElement("Started", startedTime),
                        new XElement("Ended", endedTime),
                        new XElement("Status", status)));
                    element.Save("Testing-Report.xml");
                }

                //Update "total-time" attribute
                UpdateTotalTime(DateTime.Parse(startedTime), DateTime.Parse(endedTime));
      
                //Delete Temp.xml since its of no use now
                File.Delete("Temp.xml");

                // Start Processing HTML file from here..
                GenerateHTMLFile();

            }
            else
            {
                // Since file doesn't exist and it may have been already deleted only in case of failed status,
                // Increment "failures" attribute value
                XDocument doc_Main = XDocument.Load("Testing-Report.xml");
                XElement element = doc_Main.Element("Test-Suite");


                IncrementTotalANDFailedAttributes(element);
                //Console.WriteLine("Temp.xml File doesn't exists. Hence test is Failed.");

                //Save element
                element.Save("Testing-Report.xml");

                //Generate HTML Again
                GenerateHTMLFile();

                /*
                int valFailures = Convert.ToInt32(element.Attribute("failures").Value);
                valFailures = valFailures + 1;
                element.Attribute("failures").Value = valFailures.ToString();
                

                // Update "passed" attributes values
                int valtotal = Convert.ToInt32(element.Attribute("total").Value);
                int valPassed = valtotal - valFailure;
                element.Attribute("passed").Value = valPassed.ToString();
                element.Save("Testing-Report.xml");
                */

            }
        }

        private static void UpdateTotalTime(DateTime startTime, DateTime endTime)
        {
            TimeSpan diff = endTime - startTime;

            XDocument doc = XDocument.Load("Testing-Report.xml");
            XElement element = doc.Element("Test-Suite");

            TimeSpan totalTime = TimeSpan.Parse(element.Attribute("total-time").Value) + diff;
            element.Attribute("total-time").Value = totalTime.ToString();
            element.Save("Testing-Report.xml");
        }

        private static void GenerateHTMLFile()
        {
            StringBuilder html = new StringBuilder();
            string input = "Testing-Report.xml", output = string.Empty;
            output = Path.ChangeExtension(input, "html");

            html.Append(EndTest.GetHTML5Header());
            html.Append(ProcessFile(input));
            html.Append(EndTest.GetHTML5Footer());

            // Save HTML to the output file
            File.WriteAllText(output, html.ToString());
        }

        private static void IncrementTotalANDFailedAttributes(XElement element)
        {
            // Increment "total" attribute value
            int valTotal = Convert.ToInt32(element.Attribute("total").Value);

            // Update "failure" attribute value
            int valFailure = Convert.ToInt32(element.Attribute("failures").Value);
            valFailure = valFailure + 1;
            element.Attribute("failures").Value = valFailure.ToString();

            // Update "passed" attribute value
            int valPassed = valTotal - valFailure;
            element.Attribute("passed").Value = valPassed.ToString();

            //Save element
            element.Save("Testing-Report.xml");
        }

        private static void IncrementTotalANDPassedAttributes(XElement element)
        {
            // Increment "total" attribute value
            int valTotal = Convert.ToInt32(element.Attribute("total").Value);
            valTotal = valTotal + 1;
            element.Attribute("total").Value = valTotal.ToString();

            // Update "passed" attribute value
            int valPassed = Convert.ToInt32(element.Attribute("passed").Value);
            valPassed = valPassed + 1;
            element.Attribute("passed").Value = valPassed.ToString();

            // Update Failures
            int valFailures = valTotal - valPassed;
            element.Attribute("failures").Value = valFailures.ToString();

            //Save element
            element.Save("Testing-Report.xml");
        }

        private static void CheckAndUpdateAttributes(XElement element)
        {

            //if (element.Attribute("name") == null)
            {
                element.Add(new XAttribute("name", "UiPath Project - Test Automation"));
            }
            //if (element.Attribute("total") == null)
            {
                element.Add(new XAttribute("total", "0"));
            }

            //if (element.Attribute("passed") == null)
            {
                element.Add(new XAttribute("passed", "0"));
            }

            //if (element.Attribute("failures") == null)
            {
                element.Add(new XAttribute("failures", "0"));
            }

            //if (element.Attribute("execution-date") == null)
            {
                element.Add(new XAttribute("execution-date", DateTime.Now.ToString()));
            }

            //if (element.Attribute("total-time") == null)
            {
                element.Add(new XAttribute("total-time", "0"));
            }
        }



        private static string GetHTML5Header()
        {
            StringBuilder header = new StringBuilder();
            header.AppendLine("<!doctype html>");
            header.AppendLine("<html lang=\"en\">");
            header.AppendLine("  <head>");
            header.AppendLine("    <meta charset=\"utf-8\">");
            header.AppendLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1, maximum-scale=1\" />"); // Set for mobile
            header.AppendLine(string.Format("<title>{0}</title>", "UiPath - Test Automation Report"));

            // Add custom scripts

            header.AppendLine("    <script>");
            // Include jQuery in the page
            header.AppendLine(Properties.Resources.jQuery);
            header.AppendLine("    </script>");


            header.AppendLine("    <script>");
            // Include Bootstrap in the page
            header.AppendLine(Properties.Resources.BootstrapJS);
            header.AppendLine("    </script>");


            header.AppendLine("    <script type=\"text/javascript\">");
            header.AppendLine("    $(document).ready(function() { ");
            header.AppendLine("        $('[data-toggle=\"tooltip\"]').tooltip({'placement': 'bottom'});");
            header.AppendLine("    });");
            header.AppendLine("    </script>");

            // Add custom styles
            header.AppendLine("    <style>");

            // Include Bootstrap CSS in the page
            header.AppendLine(Properties.Resources.BootstrapCSS);
            header.AppendLine("    .page { margin: 15px 0; }");
            header.AppendLine("    .no-bottom-margin { margin-bottom: 0; }");
            header.AppendLine("    .printed-test-result { margin-top: 15px; }");
            header.AppendLine("    .reason-text { margin-top: 15px; }");
            header.AppendLine("    .scroller { overflow: scroll; }");
            header.AppendLine("    @media print { .panel-collapse { display: block !important; } }");
            header.AppendLine("    .val { font-size: 38px; font-weight: bold; margin-top: -10px; }");
            header.AppendLine("    .stat { font-weight: 800; text-transform: uppercase; font-size: 0.85em; color: #6F6F6F; }");
            header.AppendLine("    .test-result { display: block; }");
            header.AppendLine("    .no-underline:hover { text-decoration: none; }");
            header.AppendLine("    .text-default { color: #555; }");
            header.AppendLine("    .text-default:hover { color: #000; }");
            header.AppendLine("    .info { color: #888; }");
            header.AppendLine("    </style>");
            header.AppendLine("  </head>");
            header.AppendLine("  <body>");

            return header.ToString();
        }

        private static string ProcessFile(string file)
        {
            StringBuilder html = new StringBuilder();
            XElement doc = XElement.Load(file);

            // Load summary values
            string testName = doc.Attribute("name").Value;
            int testTests = int.Parse(!string.IsNullOrEmpty(doc.Attribute("total").Value) ? doc.Attribute("total").Value : "0");
            int testPassed = int.Parse(!string.IsNullOrEmpty(doc.Attribute("passed").Value) ? doc.Attribute("passed").Value : "0");
            int testFailures = int.Parse(!string.IsNullOrEmpty(doc.Attribute("failures").Value) ? doc.Attribute("failures").Value : "0");
            //int testNotRun = int.Parse(!string.IsNullOrEmpty(doc.Attribute("not-run").Value) ? doc.Attribute("not-run").Value : "0");
            //int testInconclusive = int.Parse(!string.IsNullOrEmpty(doc.Attribute("inconclusive").Value) ? doc.Attribute("inconclusive").Value : "0");
            //int testIgnored = int.Parse(!string.IsNullOrEmpty(doc.Attribute("ignored").Value) ? doc.Attribute("ignored").Value : "0");
            //int testSkipped = int.Parse(!string.IsNullOrEmpty(doc.Attribute("skipped").Value) ? doc.Attribute("skipped").Value : "0");
            //int testInvalid = int.Parse(!string.IsNullOrEmpty(doc.Attribute("invalid").Value) ? doc.Attribute("invalid").Value : "0");
            DateTime testDate = DateTime.Parse(string.Format("{0}", doc.Attribute("execution-date").Value));
            TimeSpan totalTime = TimeSpan.Parse(!string.IsNullOrEmpty(doc.Attribute("total-time").Value) ? doc.Attribute("total-time").Value : "0");
            //string testPlatform = doc.Element("environment").Attribute("platform").Value;

            // Calculate the success rate
            decimal percentage = 0;
            if (testTests > 0)
            {
                int failures = testFailures;
                percentage = decimal.Round(decimal.Divide(failures, testTests) * 100, 1);
            }

            // Container
            html.AppendLine("<div class=\"container-fluid page\">");

            // Summary panel
            html.AppendLine("<div class=\"row\">");
            html.AppendLine("<div class=\"col-md-12\">");
            html.AppendLine("<div class=\"panel panel-default\">");
            html.AppendLine(string.Format("<div class=\"panel-heading\">Summary - <small>{0}</small></div>", testName));
            html.AppendLine("<div class=\"panel-body\">");

            html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Tests</div><div class=\"val ignore-val\">{0}</div></div>", testTests));
            html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Failures</div><div class=\"val {1}\">{0}</div></div>", testFailures, testFailures > 0 ? "text-danger" : string.Empty));
            html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Passed</div><div class=\"val {1}\">{0}</div></div>", testPassed, testPassed > 0 ? "text-success" : string.Empty));
            //html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Not Run</div><div class=\"val {1}\">{0}</div></div>", testNotRun, testNotRun > 0 ? "text-danger" : string.Empty));
            //html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Inconclusive</div><div class=\"val {1}\">{0}</div></div>", testInconclusive, testInconclusive > 0 ? "text-danger" : string.Empty));
            //html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Ignored</div><div class=\"val {1}\">{0}</div></div>", testIgnored, testIgnored > 0 ? "text-danger" : string.Empty));
            //html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Skipped</div><div class=\"val {1}\">{0}</div></div>", testSkipped, testSkipped > 0 ? "text-danger" : string.Empty));
            //html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Invalid</div><div class=\"val {1}\">{0}</div></div>", testInvalid, testInvalid > 0 ? "text-danger" : string.Empty));
            html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Execution Date</div><div class=\"val\">{0}</div></div>", testDate.ToString("d MMMM yyyy")));
            html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Execution Time</div><div class=\"val\">{0}</div></div>", totalTime));
            html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Platform</div><div class=\"val\">{0}</div></div>", "UiPath"));
            html.AppendLine(string.Format("<div class=\"col-md-2 col-sm-4 col-xs-6 text-center\"><div class=\"stat\">Success Rate</div><div class=\"val\">{0}%</div></div>", 100 - percentage));

            // End summary panel
            html.AppendLine("</div>");
            html.AppendLine("</div>");
            html.AppendLine("</div>");


            // Generate Table Data

            html.Append("<div class=\"container\">");
            html.Append("<h2>Test Execution Details</h2>");
            html.Append("<p>All Your Test Case Executions are displayed in the below table.</p>");
            html.Append("<table class=\"table\">");

            html.Append("<thead class=\"thead -light\">");
            html.Append("<tr>");
            html.Append("<th>Test Scenario</th>");
            html.Append("<th>Test Name</th>");
            html.Append("<th>Started</th>");
            html.Append("<th>Ended</th>");
            html.Append("<th >STATUS</th>");
            html.Append("</tr>");
            html.Append("</thead>");

            html.Append("<tbody>");
            XDocument xmlDoc = XDocument.Load("Testing-Report.xml");
            //Console.WriteLine(xmlDoc.Descendants("Result"));

            foreach (XElement element in xmlDoc.Descendants("Result"))
            {
                string testScenario = element.Element("TestScenario").Value;
                string testcaseName = element.Element("TestName").Value;
                string startTime = element.Element("Started").Value;
                string endTime = element.Element("Ended").Value;
                string status = element.Element("Status").Value;

                html.Append("<tr>");
                html.Append(string.Format("<td> {0} </td >", testScenario));
                html.Append(string.Format("<td> {0} </td >", testcaseName));
                html.Append(string.Format("<td> {0} </td >", startTime));
                html.Append(string.Format("<td> {0} </td >", endTime));
                html.Append(string.Format("<td> {0} </td >", status));
                html.Append("</tr>");

            }

            html.Append("</tbody>");
            html.Append("</table>");
            html.Append("</div>");

            // Process test fixtures
            //html.Append(ProcessFixtures(doc.Descendants("test-suite").Where(x => x.Attribute("type").Value == "TestFixture")));

            // End container
            html.AppendLine("</div>");
            html.AppendLine("</div>");

            return html.ToString();
        }

        private static string GetHTML5Footer()
        {
            StringBuilder footer = new StringBuilder();
            footer.AppendLine("  </body>");
            footer.AppendLine("</html>");

            return footer.ToString();
        }

    }
}