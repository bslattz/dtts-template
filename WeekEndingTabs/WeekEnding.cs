using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Excel;
using Workbook = Microsoft.Office.Tools.Excel.Workbook;
using WorksheetBase = Microsoft.Office.Tools.Excel.WorksheetBase;
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
			var configSheet = (Microsoft.Office.Interop.Excel.Worksheet) wb.Sheets["config"];
			if (configSheet == null) return;
			((DocEvents_Event) configSheet).Calculate += Refresh;
		}

		public void Refresh()
		{
			_update();
		}

		private void _update ()
		{
			var daySheets = new SheetsCollection();
			var sheetDateAddress = _wb.Names.Item("closingDate").RefersToRange.Address;
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
			//   then select the day sheets and
			//     order them to match the calculated sheetNames range and
			//     write the new dates in sheetDateAddress
			sheetTags.Join<string, Worksheet, string, SheetsCollection.Record>(
					allSheetsHosted,
					tag => tag, sheet => sheet.Name,
					(tag, sheet) => new SheetsCollection.Record {Key = tag, Sheet = sheet},
					daySheets.SheetToTagEquivalence)
				.Select((r, i) =>
				{
					r.Sheet.Name = sheetNames[i];
					return r;
				})
				.ToList();

			daySheets
				.Add(sheetTags.Join<string, Worksheet, string, SheetsCollection.Record>(
						allSheetsHosted, 
						name => name, sheet => sheet.Name, 
						(name, sheet) => new SheetsCollection.Record { Key = name, Sheet = sheet},
						daySheets.SheetToTagEquivalence)
					.Select((r, i) => { r.Sheet.Name = sheetNames[i]; return r; })
					.ToList())
				.Sort(sheetTags);

			days.Join<string, string, string, Worksheet>(daySheets.Keys,
				name => name, sheet => sheet,
				(name, sheet) => daySheets[name],
				daySheets.SheetToTagEquivalence)
				.Select((s, i) =>
				{
					s.Range[sheetDateAddress].Value2 = dates[i];
					return $"{s.Name}\t{dates[i]}\t{s.Range[sheetDateAddress].Value}";
				}).ToArray();
		}
		private static T[] ConvertRange<T>(object[,] r)
		{
			var l = r.GetUpperBound(1);
			var ret = new T[l];
			for (var i = 0; i < l; i++)
			{
				ret[i] = (T) r[1, i + 1];
			}
			return ret;
		}
		private class SheetsCollection : Dictionary<string, Worksheet>
		{
			public SheetsCollection Add (Record record)
			{
				base.Add(record.Key, record.Sheet);
				return this;
			}
			public SheetsCollection Add (List<Record> records)
			{
				records.ForEach(r => Add(r));
				return this;
			}
			public SheetsCollection Sort (IReadOnlyList<string> sortOrder, int delay = 0)
			{
				var actSht = Globals.ThisWorkbook.ActiveSheet as Excel.Worksheet;

				for (var i = sortOrder.Count - 2; i > -1; i--)
				{
					this[sortOrder[i]].Move(Before:this[sortOrder[i+1]].InnerObject);
					Thread.Sleep(delay);
				}

				actSht?.Activate();

				return this;
			}

			public new readonly IEqualityComparer<object> SheetToTagEquivalence = new SheetComparer();

			private class SheetComparer : IEqualityComparer<object>
			{
				public new bool Equals(object sheet, object day)
				{
					var sheetName = sheet as string;
					var dayName = day as string;

					Debug.Write($"{day} --> {sheet}\t");

					if (dayName == null && sheetName == null)
						return true;
					if (dayName == null || sheetName == null)
						return false;

					bool match = sheetName.IndexOf(dayName, StringComparison.Ordinal) != -1;

					Debug.WriteLine("{0}", match);

					return match;
				}

				public int GetHashCode(object obj)
				{
					return base.GetHashCode(); // force to use Equals
				}
			}

			public class Record
			{
				public string Key;
				public Worksheet Sheet;
			}
		}
	}

	public static class ExcelExtensions
	{
		public static T[] ToArray<T> (this Microsoft.Office.Interop.Excel.Range range)
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
