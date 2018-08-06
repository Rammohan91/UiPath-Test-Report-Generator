using System;
using System.Activities;
using System.Activities.Hosting;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace TestReportGenerator
{
    public class StartTest : CodeActivity
    {

        // Finished Creation of Start Activity
        private InArgument<String> statusValue = "STARTED";

        [Category("Status")]
        [RequiredArgument]
        public InArgument<String> Status
        {
            get { return this.statusValue; }
            set { this.statusValue = value; }
        }

        [Category("Test Details")]
        [RequiredArgument]
        public InArgument<String> ScenarioName { get; set; }

        [Category("Test Details")]
        public InArgument<String> TestCase { get; set; }

        public class WorkflowInstanceInfo : IWorkflowInstanceExtension

        {
            public IEnumerable<object> GetAdditionalExtensions()
            {
                yield break;
            }

            public void SetInstance(WorkflowInstanceProxy instance)
            {
                this.proxy = instance;
            }
            WorkflowInstanceProxy proxy;
            public WorkflowInstanceProxy GetProxy() { return proxy; }
        }

        protected override void CacheMetadata(CodeActivityMetadata metadata)

        {
            base.CacheMetadata(metadata);
            metadata.AddDefaultExtensionProvider<WorkflowInstanceInfo>(() => new WorkflowInstanceInfo());
        }

        protected override void Execute(CodeActivityContext context)
        {
            
            WorkflowInstanceProxy proxy = context.GetExtension<WorkflowInstanceInfo>().GetProxy();
            Activity root = proxy.WorkflowDefinition;

            string testname;

            if (TestCase.Get(context) != null)
            {
                testname = TestCase.Get(context);
            }
            else
            {
                testname = root.DisplayName;
            }

                XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
                xmlWriterSettings.Indent = true;
                xmlWriterSettings.NewLineOnAttributes = true;
                using (XmlWriter xmlWriter = XmlWriter.Create("Temp.xml", xmlWriterSettings))
                {
                    xmlWriter.WriteStartDocument();
                    xmlWriter.WriteStartElement("Test-Suite");
                    xmlWriter.WriteStartElement("Result");
                    xmlWriter.WriteElementString("TestScenario", ScenarioName.Get(context));
                    xmlWriter.WriteElementString("TestName", testname);
                    xmlWriter.WriteElementString("Started", DateTime.Now.ToString());
                    //xmlWriter.WriteElementString("Ended", "10:00:04");
                    xmlWriter.WriteElementString("Status", Status.Get(context));
                    xmlWriter.WriteEndElement();
                    xmlWriter.WriteEndElement();
                    xmlWriter.WriteEndDocument();
                    xmlWriter.Flush();
                    xmlWriter.Close();
                }
        }

    }
}
