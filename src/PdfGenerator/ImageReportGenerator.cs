using ImageProcessor;
using ImageProcessor.Imaging.Formats;
using iTextSharp.text;
using iTextSharp.text.pdf;
using PdfGenerator.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Telerik.Windows.Documents.Fixed.FormatProviders.Pdf;
using Telerik.Windows.Documents.Fixed.FormatProviders.Pdf.Export;
using Telerik.Windows.Documents.Fixed.Model;
using Telerik.Windows.Documents.Fixed.Model.Editing.Tables;

namespace PdfGenerator
{
    class ImageReportGenerator
    {

        public static void GenerateImageReport(Case Case)
        {
            BINLookup binLookup = new BINLookup();

            MemoryStream ms = new MemoryStream();

            TimeZoneInfo CurrentTimeZone = TimeZoneInfo.FindSystemTimeZoneById(Case.Client.TimeZoneID);
            string CurrentTime = DateTimeOffset.UtcNow.ToOffset(CurrentTimeZone.GetUtcOffset(DateTimeOffset.UtcNow)).ToString("ddd MMM d yyyy h:mm:ss tt zzz");
            Document document = new Document(PageSize.LETTER.Rotate(), 5f, 5f, 30f, 20f);
            PdfWriter writer = PdfWriter.GetInstance(document, ms);
            string strTemp = string.Empty;
            document.Open();
            Font fArial16 = FontFactory.GetFont("ARIAL", 16f);
            Font fArialBold16 = FontFactory.GetFont("ARIAL", 16f, Font.BOLD);
            Font fArial14 = FontFactory.GetFont("ARIAL", 12f);
            Font fArialBold14 = FontFactory.GetFont("ARIAL", 12f, Font.BOLD);
            Font fArialItalic14 = FontFactory.GetFont("ARIAL", 12f, Font.ITALIC);
            Font fArial10 = FontFactory.GetFont("ARIAL", 9f);
            Font fArialBold10 = FontFactory.GetFont("ARIAL", 9f, Font.BOLD);

            PdfPTable pTable1 = new PdfPTable(1);

           

            Table table = new Table();
           


            PdfPTable pTableCards = new PdfPTable(new float[] { 300f, 100f });
            document.Close();
            string fileName = "CardImageReport_" + Case.CaseNumber.Replace(" ", "") + "_" + TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(DateTime.Now, DateTimeKind.Utc), TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")).ToString("MMddyyyy") + ".pdf";
            using (FileStream file = new FileStream(fileName, FileMode.Create, System.IO.FileAccess.Write))
            {
                var bytes = ms.ToArray();
                file.Write(bytes, 0, bytes.Length);
                document.Close();
            }
        }

        public static void GenerateImageReport(Case Case, List<Card> Cards)
        {
            BINLookup binLookup = new BINLookup();

            MemoryStream ms = new MemoryStream();

            TimeZoneInfo CurrentTimeZone = TimeZoneInfo.FindSystemTimeZoneById(Case.Client.TimeZoneID);
            string CurrentTime = DateTimeOffset.UtcNow.ToOffset(CurrentTimeZone.GetUtcOffset(DateTimeOffset.UtcNow)).ToString("ddd MMM d yyyy h:mm:ss tt zzz");
            Document document = new Document(PageSize.LETTER.Rotate(), 5f, 5f, 30f, 20f);
            PdfWriter writer = PdfWriter.GetInstance(document, ms);
            string strTemp = string.Empty;
            document.Open();
            Font fArial16 = FontFactory.GetFont("ARIAL", 16f);
            Font fArialBold16 = FontFactory.GetFont("ARIAL", 16f, Font.BOLD);
            Font fArial14 = FontFactory.GetFont("ARIAL", 12f);
            Font fArialBold14 = FontFactory.GetFont("ARIAL", 12f, Font.BOLD);
            Font fArialItalic14 = FontFactory.GetFont("ARIAL", 12f, Font.ITALIC);
            Font fArial10 = FontFactory.GetFont("ARIAL", 9f);
            Font fArialBold10 = FontFactory.GetFont("ARIAL", 9f, Font.BOLD);

            PdfPTable pTable1 = new PdfPTable(1);

            int i = 0;

            PdfPTable pTableCards = new PdfPTable(new float[] { 300f, 100f });

            foreach (Card c in Cards)
            {

                #region Get BinBase / IAFCI Info -- used based on which data is available: priority is Provider Info, Bin Base, IAFCI

                var BinBase = c.BinBase;

                var IAFCI = new IAFCI();

                var Override = new BinBaseOverride();

                if (BinBase == null)
                {
                    BinBase = new BinBase();
                }
                if (IAFCI == null)
                {
                    IAFCI = new IAFCI();
                }
                if (Override == null)
                {
                    Override = new BinBaseOverride();
                }

                #endregion

                #region See if iteration has a new provider (each provider gets a new heading)

                if (i == 0 || i % 4 == 0)
                {
                    //Reset Values
                    pTableCards = new PdfPTable(new float[] { 175f, 180f, 200f, 200f });

                    #region Provider Header (Transaction Type, Balance Inquiry, Client ID, User ID, Case #,  Timestamp)

                    PdfPCell pCell = new PdfPCell(new Phrase(
                        (Case.ClientGroup != null ? Case.ClientGroup.Name + "\r\n" : "") +
                        c.Case.Client.Name + "\r\n" +
                        (c.Case.Client.Address1 + " " + c.Case.Client.Address2).Trim() + "\r\n" +
                        c.Case.Client.City + ", " + c.Case.Client.State + ", " + c.Case.Client.Zipcode + "\r\n"
                    , fArialBold14));
                    pCell.Padding = 0f;
                    pCell.PaddingBottom = 10f;
                    pCell.Border = 0;
                    pCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    pTable1.AddCell(pCell);

                    //provider title
                    pCell = new PdfPCell(new Phrase("Card Image Report", fArialBold16));
                    pCell.Padding = 0f;
                    pCell.PaddingBottom = 12f;
                    pCell.Border = 0;
                    pCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    pTable1.AddCell(pCell);

                    PdfPTable pTable2 = new PdfPTable(new float[] { 45f, 490f });

                    pCell = new PdfPCell(new Phrase("Case #:", fArial14));
                    pCell.Padding = 0f;
                    pCell.PaddingBottom = 5f;
                    pCell.PaddingRight = 5f;
                    pCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    pCell.Border = 0;
                    pTable2.AddCell(pCell);

                    pCell = new PdfPCell(new Phrase(c.Case.CaseNumber + "         Total Cards: " + ReportHelpers.GetCardCount(Cards), fArial14));
                    pCell.Padding = 0f;
                    pCell.PaddingBottom = 5f;
                    pCell.Border = 0;
                    pTable2.AddCell(pCell);

                    //add the list table to the main table
                    pCell = new PdfPCell(pTable2);
                    pCell.Padding = 0f;
                    pCell.Border = 0;
                    pTable1.AddCell(pCell);

                    //main content - paragraph 4 header
                    pCell = new PdfPCell(new Phrase(CurrentTime, fArial14));
                    pCell.Padding = 0f;
                    pCell.PaddingBottom = 8f;
                    pCell.Border = 0;
                    pTable1.AddCell(pCell);

                    #endregion


                    #region Issuer Details (Bin, Card Type, Card Level, Isuser Name)

                    pTable2 = new PdfPTable(new float[] { 200f, 490f });

                    //add the list table to the main table
                    pCell = new PdfPCell(pTable2);
                    pCell.Padding = 0f;
                    pCell.Border = 0;
                    pTable1.AddCell(pCell);

                    #endregion

                    //blank row
                    //pCell = new PdfPCell(new Phrase(" "));
                    //pCell.PaddingBottom = 5f;
                    //pCell.Border = 0;
                    //pTable1.AddCell(pCell);


                    #region Card Details Table Header

                    //header
                    //pCell = new PdfPCell(new Phrase("Report Details"));
                    //pCell.BorderColorBottom = BaseColor.BLACK;
                    //pCell.BorderWidthBottom = 2f;
                    //pCell.BorderWidthLeft = 0;
                    //pCell.BorderWidthRight = 0;
                    //pCell.BorderWidthTop = 0;
                    //pCell.PaddingBottom = 5f;
                    //pTable1.AddCell(pCell);


                    //cards table
                    pCell = new PdfPCell(new Phrase("Images", fArialBold14));
                    pCell.Padding = 5f;
                    pCell.PaddingBottom = 5f;
                    pCell.Border = 0;
                    pCell.BorderColorBottom = BaseColor.BLACK;
                    pCell.BorderWidthBottom = 2f;
                    pCell.BorderWidthLeft = 2f;
                    pCell.BorderWidthTop = 2f;
                    pTableCards.AddCell(pCell);

                    pCell = new PdfPCell(new Phrase("", fArialBold14));
                    pCell.Padding = 0f;
                    pCell.PaddingBottom = 5f;
                    pCell.PaddingRight = 5f;
                    pCell.Border = 0;
                    pCell.BorderColorBottom = BaseColor.BLACK;
                    pCell.BorderWidthBottom = 2f;
                    pCell.BorderWidthRight = 2f;
                    pCell.BorderWidthTop = 2f;
                    pTableCards.AddCell(pCell);

                    pCell = new PdfPCell(new Phrase("Printed Details", fArialBold14));
                    pCell.Padding = 5f;
                    pCell.PaddingBottom = 5f;
                    pCell.Border = 0;
                    pCell.BorderColorBottom = BaseColor.BLACK;
                    pCell.BorderWidthBottom = 2f;
                    pCell.BorderWidthRight = 2f;
                    pCell.BorderWidthTop = 2f;
                    pTableCards.AddCell(pCell);

                    pCell = new PdfPCell(new Phrase("Magstripe Details", fArialBold14));
                    pCell.Padding = 5f;
                    pCell.PaddingBottom = 5f;
                    pCell.Border = 0;
                    pCell.BorderColorBottom = BaseColor.BLACK;
                    pCell.BorderWidthBottom = 2f;
                    pCell.BorderWidthRight = 2f;
                    pCell.BorderWidthTop = 2f;
                    pTableCards.AddCell(pCell);

                    #endregion
                }

                #endregion

                #region Add Card to Cards Table (row)

                //Add Card
                //determine provider name
                //Determine Provider Name
                string ProviderName = "";
                if (c.Provider != null)
                {
                    ProviderName = (c.RetailerName != null ? c.RetailerName : c.Provider.Name);
                }
                else
                {
                    ProviderName = (string.IsNullOrEmpty(Override.IssuingOrg) ? null : Override.IssuingOrg) ??
                                    (string.IsNullOrEmpty(IAFCI.IssuerName) ? null : IAFCI.IssuerName) ??
                                    (string.IsNullOrEmpty(BinBase.IssuingOrg) ? null : BinBase.IssuingOrg) ??
                                    "Unavailable";
                }

                //get images
                var images = c.CardImages;
                if (images.Count > 0)
                {
                    try
                    {
                        var imageString = images[0].ImageData;
                        imageString = FormatBase64String(imageString);
                        byte[] bytes = Convert.FromBase64String(imageString);
                        using (MemoryStream inStream = new MemoryStream(bytes))
                        {
                            using (MemoryStream outStream = new MemoryStream())
                            {
                                // Initialize the ImageFactory using the overload to preserve EXIF metadata.
                                using (ImageFactory imageFactory = new ImageFactory(preserveExifData: true))
                                {
                                    // Load, resize, set the format and quality and save an image.
                                    imageFactory.Load(inStream)
                                                .Format(new PngFormat { Quality = 70 })
                                                .Save(outStream);
                                }
                                var streamImage = System.Drawing.Image.FromStream(outStream);
                                var image = iTextSharp.text.Image.GetInstance(streamImage, System.Drawing.Imaging.ImageFormat.Png);
                                image.ScaleToFit(124f, 110f);
                                image.SpacingAfter = 3f;

                                PdfPCell pCellCard = new PdfPCell(image);
                                pCellCard.BorderColorBottom = BaseColor.BLACK;
                                pCellCard.BorderWidthBottom = 1f;
                                pCellCard.BorderWidthLeft = 1f;
                                pCellCard.BorderWidthRight = 0;
                                pCellCard.BorderWidthTop = 0;
                                pCellCard.Padding = 5f;
                                pTableCards.AddCell(pCellCard);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        PdfPCell pCellCard = new PdfPCell(new Phrase("", fArial10));
                        pCellCard.BorderColorBottom = BaseColor.BLACK;
                        pCellCard.BorderWidthBottom = 1f;
                        pCellCard.BorderWidthLeft = 0;
                        pCellCard.BorderWidthRight = 1f;
                        pCellCard.BorderWidthTop = 0;
                        pCellCard.Padding = 5f;
                        pTableCards.AddCell(pCellCard);
                    }
                }
                else
                {
                    PdfPCell pCellCard = new PdfPCell(new Phrase("", fArial10));
                    pCellCard.BorderColorBottom = BaseColor.BLACK;
                    pCellCard.BorderWidthBottom = 1f;
                    pCellCard.BorderWidthLeft = 1f;
                    pCellCard.BorderWidthRight = 0;
                    pCellCard.BorderWidthTop = 0;
                    pCellCard.Padding = 5f;
                    pTableCards.AddCell(pCellCard);
                }

                if (images.Count > 1)
                {
                    try
                    {
                        var imageString2 =images[1].ImageData;
                        imageString2 = FormatBase64String(imageString2);
                        byte[] bytes = Convert.FromBase64String(imageString2);
                        using (MemoryStream inStream = new MemoryStream(bytes))
                        {
                            using (MemoryStream outStream = new MemoryStream())
                            {
                                // Initialize the ImageFactory using the overload to preserve EXIF metadata.
                                using (ImageFactory imageFactory = new ImageFactory(preserveExifData: true))
                                {
                                    // Load, resize, set the format and quality and save an image.
                                    imageFactory.Load(inStream)
                                                .Format(new PngFormat { Quality = 70 })
                                                .Save(outStream);
                                }
                                var streamImage = System.Drawing.Image.FromStream(outStream);
                                var image = iTextSharp.text.Image.GetInstance(streamImage, System.Drawing.Imaging.ImageFormat.Png);
                                image.ScaleToFit(124f, 110f);
                                image.SpacingAfter = 3f;

                                PdfPCell pCellCard = new PdfPCell(image);
                                pCellCard.BorderColorBottom = BaseColor.BLACK;
                                pCellCard.BorderWidthBottom = 1f;
                                pCellCard.BorderWidthLeft = 0;
                                pCellCard.BorderWidthRight = 1f;
                                pCellCard.BorderWidthTop = 0;
                                pCellCard.Padding = 5f;
                                pTableCards.AddCell(pCellCard);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        PdfPCell pCellCard = new PdfPCell(new Phrase("", fArial10));
                        pCellCard.BorderColorBottom = BaseColor.BLACK;
                        pCellCard.BorderWidthBottom = 1f;
                        pCellCard.BorderWidthLeft = 0;
                        pCellCard.BorderWidthRight = 1f;
                        pCellCard.BorderWidthTop = 0;
                        pCellCard.Padding = 5f;
                        pTableCards.AddCell(pCellCard);
                    }
                }
                else
                {
                    PdfPCell pCellCard = new PdfPCell(new Phrase("", fArial10));
                    pCellCard.BorderColorBottom = BaseColor.BLACK;
                    pCellCard.BorderWidthBottom = 1f;
                    pCellCard.BorderWidthLeft = 0;
                    pCellCard.BorderWidthRight = 1f;
                    pCellCard.BorderWidthTop = 0;
                    pCellCard.Padding = 5f;
                    pTableCards.AddCell(pCellCard);
                }

                Phrase cardNum = new Phrase("");

                if ((c.ClonedCardParent == null && !c.MissingPrintedInfo) || c.ClonedCardParent != null)
                {
                    cardNum.Add(new Chunk("Number: ", fArialBold10));
                    cardNum.Add(new Chunk(c.CardNumberMasked, fArial10));
                    cardNum.Add(new Chunk("\r\nName: ", fArialBold10));
                    cardNum.Add(new Chunk(c.CardName, fArial10));
                    cardNum.Add(new Chunk("\r\nExp: ", fArialBold10));
                    cardNum.Add(new Chunk(c.CardExpiration, fArial10));
                    cardNum.Add(new Chunk("\r\nIssuer: ", fArialBold10));

                    string IssuerName = ProviderName;
                    if (ProviderName.ToLower().Equals("unavailable") && c.IssuerName != null && c.IssuerName.Length > 0)
                    {
                        IssuerName = c.IssuerName;
                    }
                    cardNum.Add(new Chunk(IssuerName, fArial10));
                }

                PdfPCell pCardCell = new PdfPCell(cardNum);
                pCardCell.BorderColorBottom = BaseColor.BLACK;
                pCardCell.BorderWidthBottom = 1f;
                pCardCell.BorderWidthLeft = 0;
                pCardCell.BorderWidthRight = 1f;
                pCardCell.BorderWidthTop = 0;
                pCardCell.Padding = 5f;
                pCardCell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                pTableCards.AddCell(pCardCell);

                cardNum = new Phrase("");
                if (c.ClonedCardParent != null)
                {
                    if (c.ClonedCardParent.TrackData != null)
                    {
                        string cardTrackUnenc = c.ClonedCardParent.TrackData.Trim();
                        CCTrackParser trackParser = new CCTrackParser();
                        var trackData = trackParser.Parse(cardTrackUnenc);

                        if (trackData != null && trackData.TrackOne != null)
                        {
                            cardNum.Add(new Chunk("T1 Number: ", fArialBold10));
                            cardNum.Add(new Chunk(trackData.TrackOne.MaskedPAN, fArial10));
                            cardNum.Add(new Chunk("\r\nT1 Name: ", fArialBold10));
                            cardNum.Add(new Chunk(trackData.TrackOne.CardHolderName, fArial10));
                            cardNum.Add(new Chunk("\r\nT1 Exp: ", fArialBold10));
                            cardNum.Add(new Chunk(trackData.TrackOne.ExpirationDate, fArial10));
                            cardNum.Add(new Chunk("\r\nT1 Issuer: ", fArialBold10));

                            //lookup issuer baesd on this track bin (from PAN)
                            var bin1 = binLookup.LookupCardMetaData(c.BinBase);
                            cardNum.Add(new Chunk(bin1.IssuerName, fArial10));
                        }
                        else
                        {
                            cardNum.Add(new Chunk("T1 Number: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT1 Name: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT1 Exp: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT1 Issuer: ", fArialBold10));
                        }
                        //space
                        cardNum.Add(new Chunk("\r\n", fArial10));

                        if (trackData != null && trackData.TrackTwo != null)
                        {

                            cardNum.Add(new Chunk("\r\nT2 Number: ", fArialBold10));
                            cardNum.Add(new Chunk(trackData.TrackTwo.MaskedPAN, fArial10));
                            cardNum.Add(new Chunk("\r\nT2 Name: ", fArialBold10));
                            cardNum.Add(new Chunk("", fArial10));
                            cardNum.Add(new Chunk("\r\nT2 Exp: ", fArialBold10));
                            cardNum.Add(new Chunk(trackData.TrackTwo.ExpirationDate, fArial10));
                            cardNum.Add(new Chunk("\r\nT2 Issuer: ", fArialBold10));

                            //lookup issuer baesd on this track bin (from PAN)
                            var bin2 = binLookup.LookupCardMetaData(c.BinBase);
                            cardNum.Add(new Chunk(bin2.IssuerName, fArial10));
                        }
                        else
                        {
                            cardNum.Add(new Chunk("\r\nT2 Number: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT2 Name: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT2 Exp: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT2 Issuer: ", fArialBold10));
                        }
                    }
                    else
                    {
                        cardNum.Add(new Chunk("Number: ", fArialBold10));
                        cardNum.Add(new Chunk(c.ClonedCardParent.CardNumberMasked, fArial10));
                        cardNum.Add(new Chunk("\r\nName: ", fArialBold10));
                        cardNum.Add(new Chunk(c.ClonedCardParent.CardName, fArial10));
                        cardNum.Add(new Chunk("\r\nExp: ", fArialBold10));
                        cardNum.Add(new Chunk(c.ClonedCardParent.CardExpiration, fArial10));
                        cardNum.Add(new Chunk("\r\nIssuer: ", fArialBold10));
                        cardNum.Add(new Chunk(c.ClonedCardParent.IssuerName, fArial10));
                    }
                }
                else if (c.EntryMethod == "Swipe" || !c.MissingPrintedInfo)
                {
                    if (c.TrackData != null)
                    {
                        string cardTrackUnenc = c.TrackData.Trim();
                        CCTrackParser trackParser = new CCTrackParser();
                        var trackData = trackParser.Parse(cardTrackUnenc);

                        if (trackData != null && trackData.TrackOne != null)
                        {
                            cardNum.Add(new Chunk("T1 Number: ", fArialBold10));
                            cardNum.Add(new Chunk(trackData.TrackOne.MaskedPAN, fArial10));
                            cardNum.Add(new Chunk("\r\nT1 Name: ", fArialBold10));
                            cardNum.Add(new Chunk(trackData.TrackOne.CardHolderName, fArial10));
                            cardNum.Add(new Chunk("\r\nT1 Exp: ", fArialBold10));
                            cardNum.Add(new Chunk(trackData.TrackOne.ExpirationDate, fArial10));
                            cardNum.Add(new Chunk("\r\nT1 Issuer: ", fArialBold10));

                            //lookup issuer baesd on this track bin (from PAN)
                            var bin1 = binLookup.LookupCardMetaData(c.BinBase);
                            cardNum.Add(new Chunk(bin1.IssuerName, fArial10));
                        }
                        else
                        {
                            cardNum.Add(new Chunk("T1 Number: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT1 Name: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT1 Exp: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT1 Issuer: ", fArialBold10));
                        }

                        if (trackData != null && trackData.TrackTwo != null)
                        {
                            //space
                            cardNum.Add(new Chunk("\r\n", fArial10));
                            cardNum.Add(new Chunk("\r\nT2 Number: ", fArialBold10));
                            cardNum.Add(new Chunk(trackData.TrackTwo.MaskedPAN, fArial10));
                            cardNum.Add(new Chunk("\r\nT2 Name: ", fArialBold10));
                            cardNum.Add(new Chunk("", fArial10));
                            cardNum.Add(new Chunk("\r\nT2 Exp: ", fArialBold10));
                            cardNum.Add(new Chunk(trackData.TrackTwo.ExpirationDate, fArial10));
                            cardNum.Add(new Chunk("\r\nT2 Issuer: ", fArialBold10));

                            //lookup issuer baesd on this track bin (from PAN)
                            var bin2 = binLookup.LookupCardMetaData(c.BinBase);
                            cardNum.Add(new Chunk(bin2.IssuerName, fArial10));
                        }
                        else
                        {
                            cardNum.Add(new Chunk("T2 Number: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT2 Name: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT2 Exp: ", fArialBold10));
                            cardNum.Add(new Chunk("\r\nT2 Issuer: ", fArialBold10));
                        }
                    }
                    else
                    {
                        cardNum.Add(new Chunk("Number: ", fArialBold10));
                        cardNum.Add(new Chunk(c.CardNumberMasked, fArial10));
                        cardNum.Add(new Chunk("\r\nName: ", fArialBold10));
                        cardNum.Add(new Chunk(c.CardName, fArial10));
                        cardNum.Add(new Chunk("\r\nExp: ", fArialBold10));
                        cardNum.Add(new Chunk(c.CardExpiration, fArial10));
                        cardNum.Add(new Chunk("\r\nIssuer: ", fArialBold10));
                        cardNum.Add(new Chunk(ProviderName, fArial10));
                    }
                }

                pCardCell = new PdfPCell(cardNum);
                pCardCell.BorderColorBottom = BaseColor.BLACK;
                pCardCell.BorderWidthBottom = 1f;
                pCardCell.BorderWidthLeft = 0;
                pCardCell.BorderWidthRight = 1f;
                pCardCell.BorderWidthTop = 0;
                pCardCell.Padding = 5f;
                pCardCell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                pTableCards.AddCell(pCardCell);

                #endregion

                i++;

                if (i > 0 && i % 4 == 0)
                {
                    ////blank row
                    PdfPCell pCellClosing = new PdfPCell(new Phrase(""));

                    //add cards table to document
                    pCellClosing = new PdfPCell(pTableCards);
                    pCellClosing.Padding = 0f;
                    pCellClosing.Border = 0;
                    pTable1.AddCell(pCellClosing);

                    //Insert Page Break
                    document.Add(pTable1);
                    document.NewPage();
                    pTable1 = new PdfPTable(1);
                }
            }

            #region Close out Cards Table

            if (i > 0 && i % 4 != 0)
            {

                ////blank row
                PdfPCell pCellClosing = new PdfPCell(new Phrase(""));

                //add cards table to document
                pCellClosing = new PdfPCell(pTableCards);
                pCellClosing.Padding = 0f;
                pCellClosing.Border = 0;
                pTable1.AddCell(pCellClosing);

                //Insert Page Break
                document.Add(pTable1);
                document.NewPage();
                pTable1 = new PdfPTable(1);
            }

            #endregion

            document.Close();
            string fileName = "CardImageReport_" + Case.CaseNumber.Replace(" ", "") + "_" + TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(DateTime.Now, DateTimeKind.Utc), TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")).ToString("MMddyyyy") + ".pdf";
            using (FileStream file = new FileStream(fileName, FileMode.Create, System.IO.FileAccess.Write))
            {
                var bytes = ms.ToArray();
                file.Write(bytes, 0, bytes.Length);
                document.Close();
            }
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
