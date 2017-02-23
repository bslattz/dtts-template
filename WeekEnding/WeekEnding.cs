using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms.VisualStyles;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Worksheet = Microsoft.Office.Tools.Excel.Worksheet;
using Factory = Microsoft.Office.Tools.Excel.Factory;
using Excel = Microsoft.Office.Interop.Excel;

namespace WeekEnding
{
	public class WeekEnding
	{
		private readonly Microsoft.Office.Tools.Excel.WorkbookBase _wb;
		private readonly Factory _factory;
		private bool _guard;
		private DateTime[] _dates;
		private Dictionary<string, Worksheet> _taggedSheets;
	    private readonly Excel.Worksheet _configSheet;

        public WeekEnding(Microsoft.Office.Tools.Excel.WorkbookBase wb, Factory factory)
		{
			_wb = wb;
			_factory = factory;
			_configSheet = (Excel.Worksheet) wb.Sheets["config"];
            Log("WeekEnding: " + _configSheet?.Name ?? "configSheet not set");
			if (_configSheet == null) return;
			((DocEvents_Event) _configSheet).Calculate += Refresh;
		    _wb.Open += Refresh;
		}
	    ~WeekEnding()
	    {
	        Log("WeekEnding: Shutdown");
	    }
	    public void Log(string message)
	    {
	        try
	        {
	            _wb.Application.Run("Log", message);
	        }
	        catch (Exception e)
	        {
	            Console.WriteLine(">>>>>>>>>>>>>>{0}",e);
	        }
	    }
		public void Refresh()
		{
			var updating = _wb.Application.ScreenUpdating;
			_wb.Application.ScreenUpdating = false;
            _update();
			if(updating)
				_wb.Application.ScreenUpdating = true;
		}
		private void _update ()
		{
			if (_guard) return;
			var closingDateRange = _wb.Names.Item("closingDate").RefersToRange;
			var sheetDateAddress = closingDateRange.Address;
			var days = _wb. Names.Item("daysOfWeek").RefersToRange.ToArray<string>();
			_dates =_wb. Names.Item("dayDates").RefersToRange.ToArray<DateTime>();
			var sheetTags = _wb. Names.Item("datedSheets").RefersToRange.ToArray<string>();
			var sheetNames = _wb. Names.Item("datedSheetsFmt").RefersToRange.ToArray<string>();

			var allSheetsHosted = _wb.Worksheets
				.Cast<Excel.Worksheet>()
				.Select(s => _factory.GetVstoObject(s))
				.ToArray();

			var allSheets = _wb.Sheets.Cast<object>().ToArray();
			if (allSheets.Length == 0) throw new InvalidRangeException("Expecting sheets named Mon, ..., Sun, Summary");

			// Select sheets to be labeled with date info from all sheets using sheetTags
			//   then update the Sheet names to be <sheetTag><date info>
			//   and store the results in a dictionary
			_taggedSheets = sheetTags.Join(
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

			var actSht = _wb.ActiveSheet as Excel.Worksheet;

			// Select the day sheets and
			//   order them to match the calculated sheetNames range and
			//   write the new dates in sheetDateAddress
			foreach (var s in 
				days.Join<string, string, string, Worksheet>(_taggedSheets.Keys,
					name => name, sheet => sheet,
					(name, sheet) => _taggedSheets[name],
					_sheetToTagEquivalence)
				.Select((s, i) =>
				{
					if(i > 0)
						_taggedSheets[days[i]].Move(After: _taggedSheets[days[i - 1]].InnerObject);
					s.Range[sheetDateAddress].Value2 = _dates[i];
					return _dates[i];
				})
			)
			// Update the closing date on the summary sheet in case the sheets were re-ordered
			_guard = true;
			closingDateRange.Value2 = _dates.Last();
			_guard = false;

			actSht?.Activate();
		}

		public string DisplayTaggedSheets ()
		{
            if (_taggedSheets == null) return null;
            var msg = new StringBuilder();
			_taggedSheets.Aggregate(msg,
				(m, r) => m.AppendLine($"{r.Key}\t{r.Value.Name}"));
			return msg.ToString();
		}

		public string DisplayDates ()
		{
            if (_dates == null) return null;
            return string.Join("\n", _dates.Select(d => d.ToString("dd/mm/yyy")));
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
