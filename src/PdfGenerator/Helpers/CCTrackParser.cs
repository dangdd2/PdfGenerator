using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfGenerator.Helpers
{
	//https://github.com/AlexanderAnishchik/Credit-Card-Stripe-Parser
	public class CCTrackParser
	{
		#region const
		private const char _SS1 = '%';
		private const char _FS1 = '^';
		private const char _ES1 = '?';
		private const char _SS2 = ';';
		private const char _FS2 = '=';
		private const char _ES2 = '?';
		#endregion

		public FullTrackDataModel Parse(string fullTrack)
		{
			TrackOneModel track1 = null;
			TrackTwoModel track2 = null;
			bool isTrackOneValid = false;
			bool isTrackTwoValid = false;

			try
			{
				isTrackOneValid = _ValidateTrackOne(fullTrack);
				track1 = isTrackOneValid ? ParseTrackOne(fullTrack) : null;
			}
			catch (Exception)
			{
			}

			try
			{
				isTrackTwoValid = _ValidateTrackTwo(fullTrack);
				track2 = isTrackTwoValid ? ParseTrackTwo(fullTrack) : null;
			}
			catch (Exception)
			{
			}

			return new FullTrackDataModel
			{
				TrackOne = track1,
				TrackTwo = track2,
				IsTrackOneValid = isTrackOneValid,
				IsTrackTwoValid = isTrackTwoValid
			};

		}
		public TrackOneModel ParseTrackOne(string fullTrack)
		{
			string trackString = fullTrack.Substring(1, fullTrack.IndexOf(_ES1) - 1);
			if (trackString.Length > 79) throw new Exception("Track data exceeds maximum length");
			string[] trackSegments = trackString.Split(_FS1);

            string maskedPAN = MaskPAN(trackSegments[0].Substring(1));

            return new TrackOneModel
			{
				FormatCode = trackString[0],
                MaskedPAN = maskedPAN,
				PAN = trackSegments[0].Substring(1),
				CardHolderName = trackSegments[1],
				ExpirationDate = new string(trackSegments[2].Take(4).ToArray()),
				ServiceCode = new string(trackSegments[2].Skip(4).Take(3).ToArray()),
				DiscretionaryData = new string(trackSegments[2].Skip(7).ToArray()),
				SourceString = fullTrack.Substring(0, fullTrack.IndexOf(_ES1) + 1)
			};
		}
		public bool TryParseTrackOne(string fullTrack, out TrackOneModel trackOne)
		{
			try
			{
				if (!fullTrack.Contains(_SS1))
				{
					trackOne = null;
					return false;
				}
				trackOne = ParseTrackOne(fullTrack);
				return true;
			}
			catch (Exception)
			{
				trackOne = null;
				return false;
			}
		}
		public TrackTwoModel ParseTrackTwo(string fullTrack)
		{

			string trackString = fullTrack.Substring(fullTrack.IndexOf(_SS2) + 1, fullTrack.LastIndexOf(_ES2) - fullTrack.IndexOf(_SS2) - 1);
			if (trackString.Length > 40) throw new Exception("Track data exceeds maximum length");
			string[] trackSegments = trackString.Split(_FS2);
            string maskedPAN = MaskPAN(trackSegments[0]);

            return new TrackTwoModel
			{
				PAN = trackSegments[0],
                MaskedPAN = maskedPAN,
				ExpirationDate = new string(trackSegments[1].Take(4).ToArray()),
				ServiceCode = new string(trackSegments[1].Skip(4).Take(3).ToArray()),
				DiscretionaryData = new string(trackSegments[1].Skip(7).ToArray()),
				SourceString = fullTrack.Substring(fullTrack.IndexOf(_SS2), fullTrack.LastIndexOf(_ES2) - fullTrack.IndexOf(_SS2) + 1)
			};
		}
		public bool TryParseTrackTwo(string fullTrack, out TrackTwoModel trackTwo)
		{
			try
			{
				if (!fullTrack.Contains(_SS2))
				{
					trackTwo = null;
					return false;
				}
				trackTwo = ParseTrackTwo(fullTrack);
				return true;
			}
			catch (Exception)
			{
				trackTwo = null;
				return false;
			}
		}

		#region private_methods
		private byte _CalculateLRC(byte[] bytes)
		{
			return bytes.Aggregate<byte, byte>(0, (x, y) => (byte)(x ^ y));
		}
		private bool _HasLRCCode(string fullTrack)
		{
			if (fullTrack.Contains("?;") || fullTrack.EndsWith("?"))
				return false;
			return true;
		}
		private bool _ValidateTrackOne(string fullTrack)
		{
			if (!fullTrack.Contains(_SS1)) return false;

			var es1Index = fullTrack.IndexOf(_ES1);
			if (es1Index == fullTrack.Length - 1)
				return true;

			var potentialLRC = fullTrack[es1Index + 1];
			if (potentialLRC != _SS2)
			{
				var lrc = potentialLRC;
				var calculatedLRC = _CalculateLRC(fullTrack.Substring(1, es1Index).Select(c => (byte)c).ToArray());
				if (lrc != calculatedLRC)
					return false;
			}

			return true;
		}

		private bool _ValidateTrackTwo(string fullTrack)
		{
			if (!fullTrack.Contains(_SS2)) return false;
			var potentialLRC = fullTrack.Last();
			if (potentialLRC != _ES2)
			{
				var lrc = potentialLRC;
				var calculatedLRC = _CalculateLRC(fullTrack.Substring(fullTrack.IndexOf(_SS2) + 1, fullTrack.LastIndexOf(_ES2) - fullTrack.IndexOf(_SS2)).Select(c => (byte)c).ToArray());
				if (lrc != calculatedLRC)
					return false;
			}
			return true;
		}

        private string MaskPAN(string PAN)
        {
            var maskedPAN = PAN;
            //Mask Card Numbers if length is greater than 10
            if (PAN.Length > 10)
            {
                maskedPAN = PAN.Substring(0, 6) +
                                        new string('x', PAN.Length - 10) +
                                        PAN.Substring(PAN.Length - 4, 4);
            }
            return maskedPAN;
        }

		#endregion
	}


	public class FullTrackDataModel
	{
		public bool IsTrackOneValid { get; set; }
		public TrackOneModel TrackOne { get; set; }
		public bool IsTrackTwoValid { get; set; }
		public TrackTwoModel TrackTwo { get; set; }
	}
	/// <summary>
	/// ISO 7811-2 track one character encoding definition:
	/// SS FC PAN FS Name FS Date Discretionary Data ES LRC
	/// </summary>
	public class TrackOneModel
	{

		public char FormatCode { get; set; }
		public string PAN { get; set; }
		public string CardHolderName { get; set; }
		public string ExpirationDate { get; set; }
		public string ServiceCode { get; set; }
		public string DiscretionaryData { get; set; }
		public string SourceString { get; set; }
        public string MaskedPAN { get; set; }
	}
	/// <summary>
	/// ISO 7811-2 Track Two encoding definition:
	/// SS PAN FS Date SVC CD Discretionary Data ES LRC
	/// </summary>
	public class TrackTwoModel
	{
        public string MaskedPAN { get; set; }
        public string PAN { get; set; }
		public string ExpirationDate { get; set; }
		public string ServiceCode { get; set; }
		public string DiscretionaryData { get; set; }
		public string SourceString { get; set; }

	}
}
