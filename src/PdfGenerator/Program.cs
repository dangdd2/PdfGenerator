using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;

namespace PdfGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
                
            using (StreamReader r = new StreamReader("SampleCase.json"))
            {
                string json = r.ReadToEnd();
                var jObj = JsonConvert.DeserializeObject<JObject>(json);
                var Case = jObj.ToObject<Case>();

                Console.WriteLine("Generate PDF Image Report");
                PDFGenerator.GenerateImageReport(Case);

                //Console.WriteLine("Generate PDF Issuer Report");
                //IssuerReportGenerator.GenerateIssuerReport(Case, Case.Cards);
            }
        }
    }
}
