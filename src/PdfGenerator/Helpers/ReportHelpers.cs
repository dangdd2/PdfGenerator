using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PdfGenerator.Helpers
{
    class ReportHelpers
    {
        public static int GetCardCount(IEnumerable<Card> cards)
        {
            var cardsCount = cards.Count();

            return cardsCount;
        }
    }
}
