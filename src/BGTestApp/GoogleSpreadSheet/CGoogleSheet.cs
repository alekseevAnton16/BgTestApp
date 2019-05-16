using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace BGTestApp.GoogleSpreadSheet
{
	public static class CGoogleSheet
	{
		public static IList<IList<object>> GetSheetRows(SheetsService sheetsService, string range, string spreadSheetId, out bool isSuccess)
		{
			try
			{
				var request = sheetsService.Spreadsheets.Values.Get(spreadSheetId, range);
				isSuccess = true;
				return request.Execute().Values;
			}
			catch (Exception ex)
			{
				var message = $"{nameof(GetSheetRows)}: {ex.Message}";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				isSuccess = false;
				return null;
			}
		}

		/// <summary>
		/// Проверяет существующие листы таблицы и создает новые при необходимости.
		/// </summary>
		public static bool CheckAndCreateSheets(SheetsService sheetsService, string spreadSheetId, IList<Sheet> allSheets, IReadOnlyCollection<CPostgreServer> postgreServers)
		{
			if (sheetsService == null || string.IsNullOrWhiteSpace(spreadSheetId) || allSheets == null || postgreServers == null)
			{
				var message = $"{nameof(CheckAndCreateSheets)}: incorrect input parameters";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				return false;
			}

			var needUpdateSheets = false;
			var namesOfServers = postgreServers.Select(x => x.ServerName).ToList();
			var titlesOfSheets = allSheets.Select(x => x.Properties.Title).ToList();
			foreach (var serverName in namesOfServers)
			{
				if (!titlesOfSheets.Contains(serverName))
				{
					CreateSheet(sheetsService, spreadSheetId, serverName);
					needUpdateSheets = true;
				}
			}

			return needUpdateSheets;
		}


		/// <summary>
		/// Получает все листы из таблицы.
		/// </summary>
		public static IList<Sheet> GetAllSheets(SheetsService sheetsService, string spreadSheetId)
		{
			try
			{
				var spreadSheet = sheetsService?.Spreadsheets.Get(spreadSheetId).Execute();
				return spreadSheet?.Sheets;
			}
			catch (Exception ex)
			{
				var message = $"{nameof(GetAllSheets)}: {ex.Message}";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				return null;
			}
		}

		/// <summary>
		/// Создает лист в таблице.
		/// </summary>
		private static void CreateSheet(SheetsService sheetsService, string spreadSheetId, string sheetName)
		{
			var addSheetRequest = new AddSheetRequest {Properties = new SheetProperties {Title = sheetName}};
			BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
			{
				Requests = new List<Request> {new Request {AddSheet = addSheetRequest}}
			};

			var batchUpdateRequest =
				sheetsService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadSheetId);

			batchUpdateRequest.Execute();
		}
	}
}
