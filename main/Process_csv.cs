﻿using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Text;

namespace main
{
    class CSV_Record
    {
        [Name("InvoiceNumber")]
        public long InvoiceNumber { get; set; }
        [Name("DeliveryNoteNumber")]
        public string DeliveryNoteNumber { get; set; }
        [Name("CompanyCode")]
        public string CompanyCode { get; set; }
        [Name("SalesOrg")]
        public string SalesOrg { get; set; }
        [Name("Division")]
        public int Division { get; set; }
        [Name("SoldToNumber")]
        public string SoldToNumber{ get; set; }
        [Name("SoldToName")]
        public string SoldToName { get; set; }
        [Name("SoldToComMail")]
        public string SoldToComMail { get; set; }
        [Name("SoldToLanguage")]
        public string SoldToLanguage { get; set; }
        [Name("SoldToStatus")]
        public string SoldToStatus { get; set; }
        [Name("InvoiceIssueDate")]
        public string InvoiceIssueDate { get; set; }
        [Name("DocumentType")]
        public string DocumentType { get; set; }
        [Name("ShipToNumber")]
        public string ShipToNumber { get; set; }
        [Name("ShipToName")]
        public string ShipToName { get; set; }
        [Name("DeliveryNoteDate")]
        public string DeliveryNoteDate { get; set; }
        [Name("ShippingPoint")]
        public string ShippingPoint { get; set; }
    }

    class Process_csv
    {
        static void Main(string[] args)
        {
            if (!File.Exists(args[0]))
            {
                Console.WriteLine("File not found");
                return;
            }

            var config = new CsvConfiguration(CultureInfo.CurrentCulture) { Delimiter = ";", Encoding = Encoding.UTF8 };

            using (var reader = new StreamReader(args[0]))
            using (var csv = new CsvReader(reader, config))
            {
                var records = csv.GetRecords<CSV_Record>();
                string path = args[0].Remove(args[0].LastIndexOf('\\'));
                path = path.Remove(path.LastIndexOf('\\'));
                foreach (var rec in records)
                {
                    string name = rec.CompanyCode.Substring(0, 2);
                    if (!File.Exists($"{path}\\{name}_pdf\\{rec.InvoiceNumber}.pdf"))
                    {
                        if(!Directory.Exists($"{path}\\{name}_pdf")){ Directory.CreateDirectory($"{path}\\{name}_pdf"); }

                        using (FileStream fs = new FileStream($"{path}\\{name}_pdf\\{rec.InvoiceNumber}.pdf", FileMode.Create, FileAccess.Write))
                        {
                            fs.SetLength(500000);
                            Document doc = new Document(PageSize.A4);
                            PdfWriter.GetInstance(doc, fs);

                            doc.Open();

                            doc.Add(new Paragraph($"{nameof(rec.InvoiceNumber)}:{rec.InvoiceNumber}"));

                            doc.Close();
                        }
                    }
                    if (!File.Exists($"{path}\\{name}_xml\\{name}.xml"))
                    {
                        if (!Directory.Exists($"{path}\\{name}_xml")) { Directory.CreateDirectory($"{path}\\{name}_xml"); }

                        using (FileStream fs = new FileStream($"{path}\\{name}_xml\\{rec.InvoiceNumber}.xml", FileMode.Create, FileAccess.Write))
                        {
                            fs.SetLength(500000);
                            string text = $"{nameof(rec.InvoiceNumber)}:{rec.InvoiceNumber}";
                            fs.Write(Encoding.UTF8.GetBytes(text));
                        }
                    }
                }
                    
            }
            Console.WriteLine("Program has ended");
        }
    }
}
