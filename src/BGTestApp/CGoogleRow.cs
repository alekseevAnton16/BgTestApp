using System;
using System.Collections.Generic;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace BGTestApp
{
	public static class CGoogleRow
	{
		private static readonly List<string> Headers = new List<string> {"Сервер", "База данных", "Размер в ГБ", "Дата обновления"}; 

		/// <summary>
		/// Создает строку с заголовками.
		/// </summary>
		public static bool CreateHeaderRow(SheetsService sheetsService, string spreadSheetId, int? sheetId)
		{
			var requests = new List<Request>();
			for (var i = 0; i < Headers.Count; i++)
			{
				var header = Headers[i];
				var value = new CellData
				{
					UserEnteredValue = new ExtendedValue
					{
						StringValue = header
					}
				};

				requests.Add(new Request
				{
					UpdateCells = new UpdateCellsRequest
					{
						Start = new GridCoordinate
						{
							SheetId = sheetId,
							ColumnIndex = i,
						},
						Rows = new List<RowData> {new RowData {Values = new List<CellData> {value}}},
						Fields = "userEnteredValue"
					}
				});
			}

			var batchUpdate = new BatchUpdateSpreadsheetRequest
			{
				Requests = requests
			};

			try
			{
				sheetsService.Spreadsheets.BatchUpdate(batchUpdate, spreadSheetId).Execute();
				return true;
			}
			catch (Exception ex)
			{
				var message = $"{nameof(CreateHeaderRow)}: {ex.Message}";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				return false;
			}
		}
	}
}
