using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using Telerik.Documents.Core.Fonts;
using Telerik.Windows.Documents.Fixed.FormatProviders.Pdf;
using Telerik.Windows.Documents.Fixed.FormatProviders.Pdf.Streaming;
using Telerik.Windows.Documents.Fixed.Model;
using Telerik.Windows.Documents.Fixed.Model.Editing;
using Telerik.Windows.Documents.Fixed.Model.Editing.Flow;
using Telerik.Windows.Documents.Fixed.Model.Editing.Tables;
using Telerik.Windows.Documents.Fixed.Model.Fonts;
using System.Linq;
using PdfGenerator.Helpers;
using ImageProcessor;
using ImageProcessor.Imaging.Formats;
using Telerik.Windows.Documents.Fixed.Model.Resources;
using Telerik.Windows.Documents.Fixed.Model.Text;
using Telerik.Windows.Documents.Fixed.Model.ColorSpaces;

namespace PdfGenerator
{
	public class PDFGenerator
	{
        public static void GenerateImageReport(Case Case)
        {
            //BINLookup binLookup = new BINLookup();

            //MemoryStream ms = new MemoryStream();

            //TimeZoneInfo CurrentTimeZone = TimeZoneInfo.FindSystemTimeZoneById(Case.Client.TimeZoneID);
            //string CurrentTime = DateTimeOffset.UtcNow.ToOffset(CurrentTimeZone.GetUtcOffset(DateTimeOffset.UtcNow)).ToString("ddd MMM d yyyy h:mm:ss tt zzz");
            //Document document = new Document(PageSize.LETTER.Rotate(), 5f, 5f, 30f, 20f);
            //PdfWriter writer = PdfWriter.GetInstance(document, ms);
            //string strTemp = string.Empty;
            //document.Open();
                       
            var fontFamily = new Telerik.Documents.Core.Fonts.FontFamily("Arial");
            FontBase fArial16 = null;
            FontsRepository.TryCreateFont(fontFamily, FontStyles.Normal, FontWeights.Heavy,  out fArial16);
            
            //Font fArialBold16 = FontFactory.GetFont("ARIAL", 16f, Font.BOLD);
            //Font fArial14 = FontFactory.GetFont("ARIAL", 12f);
            //Font fArialBold14 = FontFactory.GetFont("ARIAL", 12f, Font.BOLD);
            //Font fArialItalic14 = FontFactory.GetFont("ARIAL", 12f, Font.ITALIC);
            //Font fArial10 = FontFactory.GetFont("ARIAL", 9f);
            //Font fArialBold10 = FontFactory.GetFont("ARIAL", 9f, Font.BOLD);
            //PdfPTable pTable1 = new PdfPTable(1);                       
            
            var block = new Block();
                        
            //block.TextProperties.Font = fArial16;
            block.HorizontalAlignment = HorizontalAlignment.Center;
            block.TextProperties.FontSize = 12f;            

            if (Case.ClientGroup != null)
            {
                block.InsertText(Case.ClientGroup.Name);
                block.InsertLineBreak();
            }

            block.InsertText(Case.Client.Name);
            block.InsertLineBreak();

            block.InsertText((Case.Client.Address1 + " " + Case.Client.Address2).Trim());
            block.InsertLineBreak();

            block.InsertText(Case.Client.City + ", " + Case.Client.State + ", " + Case.Client.Zipcode);
            block.InsertLineBreak(); 

            RadFixedDocument doc = new RadFixedDocument();

            RadFixedDocumentEditor editor = new RadFixedDocumentEditor(doc);

            editor.InsertBlock(block);
            //editor.InsertTable(table);

            block = new Block();
            block.InsertText("Card Image Report");
            block.HorizontalAlignment = HorizontalAlignment.Center;
            block.TextProperties.FontSize = 16f;
            editor.InsertBlock(block);
            TimeZoneInfo CurrentTimeZone = TimeZoneInfo.FindSystemTimeZoneById(Case.Client.TimeZoneID);
            string CurrentTime = DateTimeOffset.UtcNow.ToOffset(CurrentTimeZone.GetUtcOffset(DateTimeOffset.UtcNow)).ToString("ddd MMM d yyyy h:mm:ss tt zzz");
            
            foreach (var card in Case.Cards)
            {
                block = new Block();
                block.InsertText($"Case #: {card.Case.CaseNumber}  Total Cards: {ReportHelpers.GetCardCount(Case.Cards)}");
                block.InsertLineBreak();
                block.InsertText(CurrentTime);
                block.InsertLineBreak();
                editor.InsertBlock(block);

                Table table = new Table();
                table.LayoutType = TableLayoutType.FixedWidth;
                table.Margin = new Telerik.Documents.Primitives.Thickness { Left=0, Right=0};
                TextFragment text = new TextFragment();

                var border = new Border(2, RgbColors.Black);
                var lightBorder = new Border(1, RgbColors.Black);
                var row = table.Rows.AddTableRow();                
                
                block = new Block();                
                block.InsertText(new Telerik.Documents.Core.Fonts.FontFamily("Courier"), FontStyles.Normal, FontWeights.Heavy, "Images");
              
                var cell = row.Cells.AddTableCell();
                cell.Blocks.Add(block);
                cell.Borders = new TableCellBorders(border, border, border, border);

                block = new Block();
                block.InsertText("Printed Details");
                cell = row.Cells.AddTableCell();
                cell.Blocks.Add(block);
                cell.Borders = new TableCellBorders(border, border, border, border);

                block = new Block();
                block.InsertText("Magstripe Details");
                cell = row.Cells.AddTableCell();
                cell.Blocks.Add(block);
                cell.Borders = new TableCellBorders(border, border, border, border);

                row = table.Rows.AddTableRow();
                cell = row.Cells.AddTableCell();

                var imageTable = new Table();
                var imageRow = imageTable.Rows.AddTableRow();

                foreach (var image in card.CardImages)
                {
                    var imageCell = imageRow.Cells.AddTableCell();
                    block = new Block();           
                    block.SpacingBefore = 3f;
                    block.SpacingAfter = 3f;
                    
                    block.InsertImage(GetCaseImage(image.ImageData), 124f, 110f);
                    //imageCell.Borders = new TableCellBorders(lightBorder, lightBorder, lightBorder, lightBorder);
                    imageCell.Padding = new Telerik.Documents.Primitives.Thickness(6);
                    imageCell.Blocks.Add(block);          
                }

                cell.Borders = new TableCellBorders(lightBorder, lightBorder, lightBorder, lightBorder);
                cell.Padding = new Telerik.Documents.Primitives.Thickness(6);
                imageTable.LayoutType = TableLayoutType.AutoFit;
                cell.Blocks.Add(imageTable);

                cell = row.Cells.AddTableCell();
                block = new Block();
                block.HorizontalAlignment = HorizontalAlignment.Left;
                block.InsertText("Magstripe Details");
                cell.Borders = new TableCellBorders(lightBorder, lightBorder, lightBorder, lightBorder);
                cell.Blocks.Add(block);

                cell = row.Cells.AddTableCell();
                block = new Block();
                block.HorizontalAlignment = HorizontalAlignment.Left;
                block.InsertText("Magstripe Details");
                cell.Borders = new TableCellBorders(lightBorder, lightBorder, lightBorder, lightBorder);
                cell.Blocks.Add(block);

                editor.InsertTable(table);
                editor.InsertLineBreak();
            }

            PdfFormatProvider provider = new PdfFormatProvider();
          
            using (var stream = new FileStream("D:/MyPDF.pdf", FileMode.OpenOrCreate))
            {
                provider.Export(doc, stream);                 
            }
        }


        static Stream GetCaseImage(string imageString)
        {
            imageString = FormatBase64String(imageString);
            byte[] bytes = Convert.FromBase64String(imageString);
            MemoryStream outStream = new MemoryStream();
            using (MemoryStream inStream = new MemoryStream(bytes))
            {
                    // Initialize the ImageFactory using the overload to preserve EXIF metadata.
                    using (ImageFactory imageFactory = new ImageFactory(preserveExifData: true))
                    {
                        // Load, resize, set the format and quality and save an image.
                        imageFactory.Load(inStream)
                                    .Format(new JpegFormat { Quality = 70 })
                                    .Save(outStream);
                    }             
            }
            return outStream;
        }

        private static string FormatBase64String(string base64)
        {
            try
            {
                if (base64.Length > 0)
                {
                    StringBuilder sb = new StringBuilder(base64.Length + 5);
                    sb.Append(base64);
                    if (base64.Contains("%"))
                    {
                        sb = sb.Replace("%", "=");
                    }
                    var padding = base64.Length % 4;
                    if (padding == 3)
                    {
                        sb.Append("=");
                    }
                    else if (padding == 2)
                    {
                        sb.Append("==");
                    }
                    return sb.ToString();
                }
            }
            catch (Exception e)
            {
                return "";
            }
            return "";
        }
    }
}
