using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PdfGenerator
{
    class IssuerReportGenerator
    {
        public static void GenerateIssuerReport(Case Case, List<Card> Cards)
        {

            MemoryStream ms = new MemoryStream();
            if (Cards.Count > 0)
            {
                TimeZoneInfo CurrentTimeZone = TimeZoneInfo.FindSystemTimeZoneById(Case.Client.TimeZoneID);
                string CurrentTime = DateTimeOffset.UtcNow.ToOffset(CurrentTimeZone.GetUtcOffset(DateTimeOffset.UtcNow)).ToString("ddd MMM d yyyy h:mm:ss tt zzz");
                Document document = new Document(PageSize.LETTER, 5f, 5f, 40f, 40f);
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

                //var fontAwesomeIcon = BaseFont.CreateFont(Server.MapPath(Url.Content("~/Resources/fonts/fontawesome-webfont.ttf")), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                //Font fontAwe = new Font(fontAwesomeIcon, 8);


                PdfPTable pTable1 = new PdfPTable(1);

                int _priorProviderID = -1;
                string _priorProviderName = "!!!!";
                double dbBalanceTotal = 0.0;
                int i = 0;

                string strEncryptionString = "0Fnvwk0qzbn6sQ+qUbGDyz0MmpBt5e3=";
                PdfPTable pTableCards = new PdfPTable(new float[] { 150f, 200f, 70f, 160f, 120f });
                string BinList = "";

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

                    #region Determine Providers Name: priority is Provider Info, Bin Base, IAFCI

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

                    #endregion

                    #region See if iteration has a new provider (each provider gets a new heading)

                    if (ProviderName != _priorProviderName)
                    {
                        #region New Provider, close out previous table if there was one (add summary, add page break)

                        if (i > 0)
                        {
                            try
                            {
                                ((PdfPTable)pTable1.Rows[5].GetCells()[0].Table).Rows[0].GetCells()[1].Phrase[0] = new Phrase(BinList, fArial14);
                            }
                            catch (Exception) { }

                            ////blank row
                            PdfPCell pCellClosing = new PdfPCell(new Phrase(""));

                            //total row
                            pCellClosing = new PdfPCell(new Phrase("Balance Inquiry Amount:", fArialBold14));
                            pCellClosing.Padding = 0f;
                            pCellClosing.PaddingBottom = 5f;
                            pCellClosing.Border = 0;
                            pCellClosing.Colspan = 3;
                            pCellClosing.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                            pTableCards.AddCell(pCellClosing);

                            string Balance = "";
                            if (_priorProviderID == 0)
                            {
                                Balance = "N/A";
                            }
                            else
                            {
                                Balance = dbBalanceTotal.ToString("c");
                            }

                            pCellClosing = new PdfPCell(new Phrase(Balance, fArialBold14));
                            pCellClosing.Padding = 0f;
                            pCellClosing.PaddingBottom = 5f;
                            pCellClosing.Border = 0;
                            pCellClosing.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                            pTableCards.AddCell(pCellClosing);

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

                        //Reset Values
                        pTableCards = new PdfPTable(new float[] { 150f, 200f, 70f, 160f, 120f });
                        dbBalanceTotal = 0;

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
                        pCell = new PdfPCell(new Phrase(ProviderName, fArialBold16));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 15f;
                        pCell.Border = 0;
                        pCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        pTable1.AddCell(pCell);

                        PdfPTable pTable2 = new PdfPTable(new float[] { 200f, 490f });

                        pCell = new PdfPCell(new Phrase("Transaction Type:", fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.PaddingRight = 15f;
                        pCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase("Balance Inquiry", fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        //blank row
                        pCell = new PdfPCell(new Phrase(""));
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);
                        pTable2.AddCell(pCell);

                        if (c.Case.ClientGroup != null)
                        {
                            pCell = new PdfPCell(new Phrase("Group:", fArial14));
                            pCell.Padding = 0f;
                            pCell.PaddingBottom = 5f;
                            pCell.PaddingRight = 15f;
                            pCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                            pCell.Border = 0;
                            pTable2.AddCell(pCell);

                            pCell = new PdfPCell(new Phrase(c.Case.ClientGroup.Name, fArial14));
                            pCell.Padding = 0f;
                            pCell.PaddingBottom = 5f;
                            pCell.Border = 0;
                            pTable2.AddCell(pCell);
                        }

                        pCell = new PdfPCell(new Phrase("Client:", fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.PaddingRight = 15f;
                        pCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase(c.Case.Client.Login.ToString(), fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase("User:", fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 15f;
                        pCell.PaddingRight = 15f;
                        pCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase(c.Case.AssignedUser.DisplayName, fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 15f;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase("Case #:", fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 15f;
                        pCell.PaddingRight = 15f;
                        pCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase(c.Case.CaseNumber, fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 15f;
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
                        pCell.PaddingBottom = 15f;
                        pCell.Border = 0;
                        pTable1.AddCell(pCell);

                        #endregion


                        #region Issuer Details (Bin, Card Type, Card Level, Isuser Name)

                        pTable2 = new PdfPTable(new float[] { 200f, 490f });

                        //Provider Details
                        pCell = new PdfPCell(new Phrase("Issuer Details:", fArialBold14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTable1.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase("Bin Number:", fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.PaddingRight = 15f;
                        pCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        BinList = c.BinBaseBIN != null ? c.BinBaseBIN : "";
                        pCell = new PdfPCell(new Phrase(c.BinBaseBIN, fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        if (BinBase != null && BinBase.BIN != null && BinBase.BIN.Length > 0)
                        {
                            pCell = new PdfPCell(new Phrase("Card Type:", fArial14));
                            pCell.Padding = 0f;
                            pCell.PaddingBottom = 5f;
                            pCell.PaddingRight = 15f;
                            pCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                            pCell.Border = 0;
                            pTable2.AddCell(pCell);

                            pCell = new PdfPCell(new Phrase(BinBase.TypeOfCard, fArial14));
                            pCell.Padding = 0f;
                            pCell.PaddingBottom = 5f;
                            pCell.Border = 0;
                            pTable2.AddCell(pCell);

                            pCell = new PdfPCell(new Phrase("Card Level:", fArial14));
                            pCell.Padding = 0f;
                            pCell.PaddingBottom = 5f;
                            pCell.PaddingRight = 15f;
                            pCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                            pCell.Border = 0;
                            pTable2.AddCell(pCell);

                            pCell = new PdfPCell(new Phrase(BinBase.Category, fArial14));
                            pCell.Padding = 0f;
                            pCell.PaddingBottom = 5f;
                            pCell.Border = 0;
                            pTable2.AddCell(pCell);
                        }

                        pCell = new PdfPCell(new Phrase("Issuer Name:", fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.PaddingRight = 15f;
                        pCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase(ProviderName, fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase("Issuer Phone:", fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.PaddingRight = 15f;
                        pCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        string phoneNumber = "";
                        if (Override != null && Override.BIN != null && Override.BIN.Length > 0 && Override.IssuingOrgPhone != null && Override.IssuingOrgPhone.Length > 0)
                        {
                            phoneNumber = Override.IssuingOrgPhone;
                        }
                        else if (IAFCI != null && IAFCI.BIN != null && IAFCI.BIN.Length > 0 && IAFCI.PhoneNumber != null && IAFCI.PhoneNumber.Length > 0)
                        {
                            phoneNumber = IAFCI.PhoneNumber;
                        }
                        else if (BinBase != null && BinBase.BIN != null && BinBase.BIN.Length > 0 && BinBase.IssuingOrgPhone != null && BinBase.IssuingOrgPhone.Length > 0)
                        {
                            phoneNumber = BinBase.IssuingOrgPhone;
                        }
                        else if (c.Provider != null)
                        {
                            phoneNumber = (c.Provider.SubpoenaPhone != null ? "Subpoena: " + c.Provider.SubpoenaPhone : "");
                            phoneNumber += phoneNumber.Length > 0 && c.Provider.CorporatePhone != null ? ", " : "";
                            phoneNumber += c.Provider.CorporatePhone != null && c.Provider.CorporatePhone.Length > 0 ? "Corporate: " + c.Provider.CorporatePhone : "";
                        }
                        pCell = new PdfPCell(new Phrase(phoneNumber, fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        //add the list table to the main table
                        pCell = new PdfPCell(pTable2);
                        pCell.Padding = 0f;
                        pCell.Border = 0;
                        pTable1.AddCell(pCell);

                        //blank row
                        pCell = new PdfPCell(new Phrase(" "));
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTable1.AddCell(pCell);

                        #endregion


                        #region Issuer Address, Issuer Contact

                        //new table, two columns
                        pTable2 = new PdfPTable(2);
                        pTable2.WidthPercentage = 100;

                        pCell = new PdfPCell(new Phrase("Issuer Address:", fArialBold14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.PaddingRight = 5f;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase("Issuer Contact:", fArialBold14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);


                        string ProviderAddress = "N/A";
                        if (c.Provider == null && IAFCI.PhysicalAddress != null)
                        {
                            ProviderAddress = IAFCI.PhysicalAddress.Replace("<br>", "\r\n").Replace("<br/>", "\r\n").Replace("<br />", "\r\n").Replace("<BR>", "\r\n").Replace("<BR/>", "\r\n").Replace("<BR />", "\r\n");
                        }

                        pCell = new PdfPCell(new Phrase(ProviderAddress, fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.PaddingRight = 5f;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        string PrincipalContact = "N/A";
                        if (c.Provider == null && IAFCI.Address != null)
                        {
                            PrincipalContact = IAFCI.Address.Replace("<br>", "\r\n").Replace("<br/>", "\r\n").Replace("<br />", "\r\n").Replace("<BR>", "\r\n").Replace("<BR/>", "\r\n").Replace("<BR />", "\r\n");
                        }
                        else if (c.Provider != null)
                        {
                            PrincipalContact = (c.Provider.SubpoenaAttention != null ? c.Provider.SubpoenaAttention + "\r\n" : "") +
                                (c.Provider.SubpoenaName != null ? c.Provider.SubpoenaName + "\r\n" : "") +
                                (c.Provider.SubpoenaAddress1 != null ? c.Provider.SubpoenaAddress1 + "\r\n" : "") +
                               (c.Provider.SubpoenaAddress2 != null ? c.Provider.SubpoenaAddress2 + "\r\n" : "") +
                                (c.Provider.SubpoenaCity != null ? c.Provider.SubpoenaCity + ", " + c.Provider.SubpoenaState + " " + c.Provider.SubpoenaZip + "\r\n" : "");
                        }

                        pCell = new PdfPCell(new Phrase(PrincipalContact, fArial14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTable2.AddCell(pCell);

                        pCell = new PdfPCell(pTable2);
                        pCell.Padding = 0f;
                        pCell.Border = 0;
                        pTable1.AddCell(pCell);

                        //blank row
                        pCell = new PdfPCell(new Phrase(" "));
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTable1.AddCell(pCell);

                        #endregion


                        #region Card Details Table Header

                        //header
                        pCell = new PdfPCell(new Phrase("Card Details"));
                        pCell.BorderColorBottom = BaseColor.BLACK;
                        pCell.BorderWidthBottom = 1f;
                        pCell.BorderWidthLeft = 0;
                        pCell.BorderWidthRight = 0;
                        pCell.BorderWidthTop = 0;
                        pCell.PaddingBottom = 5f;
                        pTable1.AddCell(pCell);


                        //cards table
                        pCell = new PdfPCell(new Phrase("Issuer", fArialBold14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTableCards.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase("Card Number", fArialBold14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTableCards.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase("Cloned", fArialBold14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTableCards.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase("Name", fArialBold14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.Border = 0;
                        pTableCards.AddCell(pCell);

                        pCell = new PdfPCell(new Phrase("Balance", fArialBold14));
                        pCell.Padding = 0f;
                        pCell.PaddingBottom = 5f;
                        pCell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        pCell.Border = 0;
                        pTableCards.AddCell(pCell);

                        #endregion
                    }

                    #endregion

                    #region Add Card to Cards Table (row)

                    if (c.BinBaseBIN != null && c.BinBaseBIN.Length > 0 && !BinList.Contains(c.BinBaseBIN))
                    {
                        BinList += ", " + c.BinBaseBIN;
                    }

                    //Add Card
                    //card number
                    string IssuerName = ProviderName;
                    if (ProviderName.ToLower().Equals("unavailable") && c.IssuerName != null && c.IssuerName.Length > 0)
                    {
                        IssuerName = c.IssuerName;
                    }
                    PdfPCell pCellCard = new PdfPCell(new Phrase(IssuerName, fArial14));
                    pCellCard.Padding = 0f;
                    pCellCard.PaddingBottom = 5f;
                    pCellCard.Border = 0;
                    pTableCards.AddCell(pCellCard);

                    string CardNumber = "";
                    CardNumber = c.CardNumberMasked;
                    Phrase cardNum = new Phrase(CardNumber, fArial14);
                    if (c.EntryMethod != null && c.ClonedCardParentCardId == null && c.EntryMethod.Equals("Swipe"))
                    {
                        cardNum.Add(new Chunk(" [M]", fArial10));
                    }
                    else if (c.EntryMethod != null && c.EntryMethod.Equals("Upload"))
                    {
                        cardNum.Add(new Chunk(" [U]", fArial10));
                    }
                    else if (c.EntryMethod != null && c.EntryMethod.Equals("Chip"))
                    {
                        cardNum.Add(new Chunk(" [C]", fArial10));
                    }
                    else
                    {
                        cardNum.Add(new Chunk(" [K]", fArial10));
                    }
                    pCellCard = new PdfPCell(cardNum);
                    pCellCard.Padding = 0f;
                    pCellCard.PaddingBottom = 5f;
                    pCellCard.Border = 0;
                    pTableCards.AddCell(pCellCard);

                    pCellCard = new PdfPCell(new Phrase(c.ClonedCardParentCardId != null || c.ClonedCards.Count > 0 ? "[CL]" : "", fArial14));
                    pCellCard.Padding = 0f;
                    pCellCard.PaddingBottom = 5f;
                    pCellCard.Border = 0;
                    pTableCards.AddCell(pCellCard);

                    pCellCard = new PdfPCell(new Phrase(c.CardName, fArial14));
                    pCellCard.Padding = 0f;
                    pCellCard.PaddingBottom = 5f;
                    pCellCard.Border = 0;
                    pTableCards.AddCell(pCellCard);

                    double dbTemp = 0.0;

                    string Amount;
                    if (double.TryParse(c.OriginalBalance.ToString(), out dbTemp))
                    {
                        Amount = dbTemp.ToString("c");
                    }
                    else
                    {
                        Amount = "N/A";
                    }
                    dbBalanceTotal += dbTemp;

                    pCellCard = new PdfPCell(new Phrase(Amount, fArial14));
                    pCellCard.Padding = 0f;
                    pCellCard.PaddingBottom = 5f;
                    pCellCard.Border = 0;
                    pCellCard.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                    pTableCards.AddCell(pCellCard);

                    #endregion

                    //Set Current Provider
                    _priorProviderID = -1;
                    if (c.Provider != null)
                    {
                        _priorProviderID = c.ProviderID.Value;
                    }
                    _priorProviderName = ProviderName;
                    i++;
                }

                #region Close out Cards Table

                if (i > 0)
                {

                    ////blank row
                    PdfPCell pCellClosing = new PdfPCell(new Phrase(""));

                    //total row
                    pCellClosing = new PdfPCell(new Phrase("Balance Inquiry Amount:", fArialBold14));
                    pCellClosing.Padding = 0f;
                    pCellClosing.PaddingBottom = 5f;
                    pCellClosing.Border = 0;
                    pCellClosing.Colspan = 3;
                    pCellClosing.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                    pTableCards.AddCell(pCellClosing);

                    string Balance = "";
                    if (_priorProviderID == 0)
                    {
                        Balance = "N/A";
                    }
                    else
                    {
                        Balance = dbBalanceTotal.ToString("c");
                    }

                    pCellClosing = new PdfPCell(new Phrase(Balance, fArialBold14));
                    pCellClosing.Padding = 0f;
                    pCellClosing.PaddingBottom = 5f;
                    pCellClosing.Border = 0;
                    pCellClosing.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                    pTableCards.AddCell(pCellClosing);

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

                byte[] bytesInStream = ms.ToArray();

                string fileName = "CardIssuerReport_";
                fileName = fileName + Case.CaseNumber.Replace(" ", "") + "_" + TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(DateTime.Now, DateTimeKind.Utc), TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time")).ToString("MMddyyyy") + ".pdf";
                using (FileStream file = new FileStream(fileName, FileMode.Create, System.IO.FileAccess.Write))
                {
                    var bytes = ms.ToArray();
                    file.Write(bytes, 0, bytes.Length);
                    document.Close();
                }
            }
        }
    }
}
