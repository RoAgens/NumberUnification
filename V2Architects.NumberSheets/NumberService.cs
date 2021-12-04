using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace V2Architects.NumberSheets
{
    class NumberService
    {
        private readonly string _codeSymbol = "\u202A";
        private readonly string _uniqCodeSymbol = "\u202B";
        private readonly string _groupParameterName = "Группа";
        private readonly string _subgroupParameterName = "Подгруппа";
        private readonly string _numberPattern = @"(\d+([\.|\,]\d+)?)";

        private List<ViewSheet> _sheets;
        private string _newUniqCode;


        public NumberService(List<ViewSheet> sheets)
        {
            _sheets = sheets;
        }


        public void ReplacePrefixWithUnicode()
        {
            var unicodesForGroups = GetUnicodesForGroups();

            foreach (var sheet in _sheets)
            {
                if (SheetHasNotNubmer(sheet))
                {
                    continue;
                }

                var unicodeForSubgroup = unicodesForGroups[GetGroupKey(sheet)];
                var numberWithoutPrefix = GetNumberWithoutPrefix(sheet);
                sheet.SheetNumber = unicodeForSubgroup + numberWithoutPrefix;
            }
        }

        public void ReplaceUnicodeWithPrefix()
        {
            var unicodesForGroups = GetUnicodesForGroups();

            foreach (var sheet in _sheets)
            {
                if (SheetHasNotNubmer(sheet))
                {
                    continue;
                }

                var unicodeForSubgroup = unicodesForGroups[GetGroupKey(sheet)];
                var prefix = sheet.LookupParameter(_subgroupParameterName).AsString();
                var numberWithoutPrefix = GetNumberWithoutPrefix(sheet);
                sheet.SheetNumber = unicodeForSubgroup + prefix + " " + numberWithoutPrefix;
            }
        }

        private Dictionary<string, string> GetUnicodesForGroups()
        {
            var unicodeForSubgroups = new Dictionary<string, string>();
            var startCode = string.Empty;

            var subgroups = _sheets.GroupBy(s => GetGroupKey(s));
            foreach (var group in subgroups)
            {
                startCode = startCode + _codeSymbol;
                unicodeForSubgroups[group.Key] = startCode;
            }

            return unicodeForSubgroups;
        }

        private string GetGroupKey(ViewSheet sheet)
        {
            return sheet.LookupParameter(_groupParameterName).AsString() + sheet.LookupParameter(_subgroupParameterName).AsString();
        }

        private bool SheetHasNotNubmer(ViewSheet sheet)
        {
            var matchCollection = Regex.Matches(sheet.SheetNumber, _numberPattern);
            if (matchCollection.Count == 0)
            {
                return true;
            }

            return false;
        }

        private string GetNumberWithoutPrefix(ViewSheet viewSheet)
        {
            List<string> numberParts = new List<string>();
            var matchs = Regex.Matches(viewSheet.SheetNumber, _numberPattern).GetEnumerator();
            while (matchs.MoveNext())
            {
                numberParts.Add(matchs.Current.ToString());
            }

            return numberParts.Last();
        }

        public void ReplaceUnicodeWithPrefix(out List<string> wrongNubmers)
        {
            wrongNubmers = new List<string>();

            var subgroups = _sheets.GroupBy(s => s.LookupParameter(_subgroupParameterName).AsString());
            foreach (var subgroup in subgroups)
            {
                var dublicatedNumbers = subgroup.Select(s => GetNumberWithoutUnicode(s))
                                                .GroupBy(n => n)
                                                .Where(g => g.Count() > 1)
                                                .Select(g => g.Key)
                                                .ToList();

                foreach (var viewSheet in subgroup)
                {
                    var sheetNumber = GetNumberWithoutUnicode(viewSheet);

                    if (dublicatedNumbers.Contains(sheetNumber))
                    {
                        viewSheet.SheetNumber = GetUniqUnicode() + sheetNumber;
                        wrongNubmers.Add(viewSheet.SheetNumber + " - " + viewSheet.Name);
                    }

                    viewSheet.SheetNumber = subgroup.Key + " " + sheetNumber; 
                }
            }
        }

        private string GetNumberWithoutUnicode(ViewSheet viewSheet)
        {
            return viewSheet.SheetNumber.Replace(_codeSymbol, string.Empty);
        }

        private string GetUniqUnicode()
        {
            if (_newUniqCode == null)
            {
                _newUniqCode = _uniqCodeSymbol;
            }

            return _newUniqCode + _uniqCodeSymbol;
        }
    }
}
