using System;

namespace PdfGenerator.Helpers
{
	public class BINLookup
	{

		public CardMetaDataVM LookupCardMetaData(BinBase binbase)
		{
			CardMetaDataVM _c = new CardMetaDataVM();

			try
			{
				string BIN = "";
				if(binbase.BIN.Length > 6)
				{
					BIN = binbase.BIN.Substring(0, 6);
				}

				if (BIN.Length > 6)
				{
					BIN = BIN.Substring(0, 6);
				}
				BinBase _binbase = binbase;

				if (_binbase != null)
				{
					_c.Brand = _binbase.CardBrand;
					_c.TypeOfCard = _binbase.TypeOfCard;
					_c.Category = _binbase.Category;
					_c.Country = _binbase.IssuingCountryName;
					_c.Website = _binbase.IssuingOrgWebsite;
					_c.IssuerName = _binbase.IssuingOrg;
					_c.PhoneNumber = _binbase.IssuingOrgPhone;
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
			}

			return _c;
		}
	}
}