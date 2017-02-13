using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Worksheet = Microsoft.Office.Tools.Excel.Worksheet;
using Excel = Microsoft.Office.Interop.Excel;

namespace WeekEndingTabs
{
	class WeekEnding
	{
		private readonly ThisWorkbook _wb;

		public WeekEnding(ThisWorkbook wb)
		{
			_wb = wb;
			var configSheet = (Excel.Worksheet) wb.Sheets["config"];
			if (configSheet == null) return;
			((DocEvents_Event) configSheet).Calculate += Refresh;
		}
		public void Refresh()
		{
			var updating = _wb.Application.ScreenUpdating;
			_wb.Application.ScreenUpdating = false;
			_update();
			if(updating)
				_wb.Application.ScreenUpdating = true;
		}
		private bool _guard;
		private void _update ()
		{
			if (_guard) return;
			var closingDateRange = _wb.Names.Item("closingDate").RefersToRange;
			var sheetDateAddress = closingDateRange.Address;
			var days = _wb. Names.Item("daysOfWeek").RefersToRange.ToArray<string>();
			var dates =_wb. Names.Item("dayDates").RefersToRange.ToArray<DateTime>();
			var sheetTags = _wb. Names.Item("datedSheets").RefersToRange.ToArray<string>();
			var sheetNames = _wb. Names.Item("datedSheetsFmt").RefersToRange.ToArray<string>();

			var allSheetsHosted = _wb.Worksheets
				.Cast<Excel.Worksheet>()
				.Select(s => Globals.Factory.GetVstoObject(s))
				.ToArray();

			var allSheets = _wb.Sheets.Cast<object>().ToArray();
			if (allSheets.Length == 0) throw new InvalidRangeException("Expecting sheets named Mon, ..., Sun, Summary");

			// Select sheets to be labeled with date info from all sheets using sheetTags
			//   then update the Sheet names to be <sheetTag><date info>
			//   and store the results in a dictionary
			var taggedSheets = sheetTags.Join(
					allSheetsHosted,
					tag => tag, sheet => sheet.Name,
					(tag, sheet) => new {Key = tag, Sheet = sheet},
					_sheetToTagEquivalence)
				.Select((r, i) =>
				{
					r.Sheet.Name = sheetNames[i];
					return r;
				})
				.ToDictionary(r => r.Key, r => r.Sheet);

			var actSht = Globals.ThisWorkbook.ActiveSheet as Excel.Worksheet;

			// Select the day sheets and
			//   order them to match the calculated sheetNames range and
			//   write the new dates in sheetDateAddress
			foreach (var s in 
				days.Join<string, string, string, Worksheet>(taggedSheets.Keys,
					name => name, sheet => sheet,
					(name, sheet) => taggedSheets[name],
					_sheetToTagEquivalence)
				.Select((s, i) =>
				{
					if(i > 0)
						taggedSheets[days[i]].Move(After: taggedSheets[days[i - 1]].InnerObject);
					s.Range[sheetDateAddress].Value2 = dates[i];
					return dates[i];
				})
			)
			// Update the closing date on the summary sheet in case the sheets were re-ordered
			_guard = true;
			closingDateRange.Value2 = dates.Last();
			_guard = false;

			actSht?.Activate();
		}
		private readonly IEqualityComparer<object> _sheetToTagEquivalence = new SheetComparer();
		private class SheetComparer : IEqualityComparer<object>
		{
			bool IEqualityComparer<object>.Equals (object sheet, object day)
			{
				var sheetName = sheet as string;
				var dayName = day as string;

				if (dayName == null && sheetName == null)
					return true;
				if (dayName == null || sheetName == null)
					return false;

				bool match = sheetName.IndexOf(dayName, StringComparison.Ordinal) != -1;

				return match;
			}
			public int GetHashCode (object obj)
			{
				return base.GetHashCode(); // force to use Equals
			}
		}
	}

	public static class ExcelExtensions
	{
		public static T[] ToArray<T> (this Range range)
		{
			var r = range.Value2;
			var rows = r.GetUpperBound(0);
			var columns = r.GetUpperBound(1);
			var ret = new T[columns];
			for (var row = 0; row < rows; row++)
			{
				for (var col = 0; col < columns; col++)
				{
					ret[col] = (T)(typeof(T) == typeof(DateTime) ? DateTime.FromOADate(r[row + 1, col + 1]) : r[row + 1, col + 1]);
				}
				
			}
			return ret;
		}
	}
}
