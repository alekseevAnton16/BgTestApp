using System;
using System.Collections.Generic;
using System.Linq;
using BGTestApp.Enums;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace BGTestApp.GoogleSpreadSheet
{
	public static class CGoogleSheet
	{
		private const string FirstSheetNameInRussian = "Лист1";
		private const string FirstSheetNameInEnglish = "Sheet1"; 

		/// <summary>
		/// Возвращает список строк таблицы.
		/// </summary>
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
		/// Проверяет существующие листы таблицы и создает новые при необходимости.
		/// </summary>
		public static bool CheckAndCreateSheets(SheetsService sheetsService, string spreadSheetId, IList<Sheet> allSheets, IReadOnlyCollection<CPostgreServer> postgreServers)
		{
			if (sheetsService == null || string.IsNullOrWhiteSpace(spreadSheetId) || allSheets == null || !(postgreServers?.Any() ?? false))
			{
				var message = $"{nameof(CheckAndCreateSheets)}: incorrect input parameters";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				return false;
			}
			
			var updateResult = UpdateFirstSheetTitle(sheetsService, spreadSheetId, allSheets, postgreServers.First().ServerName);
			
			var namesOfServers = postgreServers.Select(x => x.ServerName).ToList();
			var titlesOfSheets = allSheets.Select(x => x.Properties.Title).ToList();
			var addResult = CreateSheets(sheetsService, spreadSheetId, namesOfServers, titlesOfSheets);

			return updateResult || addResult;
		}

		/// <summary>
		/// Создает новые листы.
		/// </summary>
		private static bool CreateSheets(SheetsService sheetsService, string spreadSheetId, List<string> namesOfServers, List<string> titlesOfSheets)
		{
			var needUpdateSheets = false;
			foreach (var serverName in namesOfServers)
			{
				if (!titlesOfSheets.Contains(serverName))
				{
					GetCreateSheetRequest(sheetsService, spreadSheetId, serverName);
					var createResult = UpdateSheet(sheetsService, spreadSheetId, EAction.Add, null, serverName);
					if (createResult)
					{
						needUpdateSheets = true;
					}
				}
			}

			return needUpdateSheets;
		}

		/// <summary>
		/// Обновление/ создание листа.
		/// </summary>
		private static bool UpdateFirstSheetTitle(SheetsService sheetsService, string spreadSheetId, IList<Sheet> allSheets, string firstSheetTitle)
		{
			if (!allSheets.Any())
			{
				return false;
			}

			var sheetInRussian = allSheets.FirstOrDefault(x => string.Equals(x.Properties.Title, FirstSheetNameInRussian));
			var sheetInEnglish = allSheets.FirstOrDefault(x => string.Equals(x.Properties.Title, FirstSheetNameInEnglish));
			var targetSheet = sheetInRussian ?? sheetInEnglish;
			return UpdateSheet(sheetsService, spreadSheetId, EAction.Update, targetSheet?.Properties?.SheetId, firstSheetTitle);
		}

		/// <summary>
		/// Создание/обновление листа.
		/// </summary>
		private static bool UpdateSheet(SheetsService sheetsService, string spreadSheetId, EAction action, int? sheetId, string sheetName)
		{
			var request = action == EAction.Add
				? GetCreateSheetRequest(sheetsService, spreadSheetId, sheetName)
				: GetUpdateSheetRequest(sheetsService, spreadSheetId, sheetId, sheetName);
			try
			{
				request.Execute();
				return true;
			}
			catch (Exception e)
			{
				var message = $"{nameof(UpdateSheet)}(action = {action}): {e.Message}";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				return false;
			}
		}

		#region Requests

		/// <summary>
		/// Получает запрос на создание листа.
		/// </summary>
		private static SpreadsheetsResource.BatchUpdateRequest GetCreateSheetRequest(SheetsService sheetsService, string spreadSheetId, string sheetName)
		{
			if (string.IsNullOrWhiteSpace(sheetName))
			{
				return null;
			}

			var addSheetRequest = new AddSheetRequest {Properties = new SheetProperties {Title = sheetName}};
			var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
			{
				Requests = new List<Request> {new Request {AddSheet = addSheetRequest}}
			};

			return sheetsService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadSheetId);
		}

		/// <summary>
		/// Получает запрос на обновление листа.
		/// </summary>
		private static SpreadsheetsResource.BatchUpdateRequest GetUpdateSheetRequest(SheetsService sheetsService, string spreadSheetId, int? sheetId, string sheetName)
		{
			if (sheetId == null || string.IsNullOrWhiteSpace(sheetName))
			{
				return null;
			}

			var updateSheet = new UpdateSheetPropertiesRequest
			{
				Properties = new SheetProperties
				{
					SheetId = sheetId,
					Title = sheetName
				},
				Fields = "title"
			};

			BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
			{
				Requests = new List<Request> {new Request {UpdateSheetProperties = updateSheet}}
			};

			return sheetsService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadSheetId);
		}

		#endregion
	}
}
