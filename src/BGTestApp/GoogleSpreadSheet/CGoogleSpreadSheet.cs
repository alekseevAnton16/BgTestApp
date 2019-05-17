using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using BGTestApp.Enums;
using BGTestApp.Properties;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace BGTestApp.GoogleSpreadSheet
{
	public class CGoogleSpreadSheet
	{
		private const string ApplicationName = "BGTestApp";
		private const string TokenFile = "token.json";
		private const string DefaultUser = "user";

		private static readonly string[] ScopesSheets = {SheetsService.Scope.Spreadsheets};

		/// <summary>
		/// Наименование создаваемой таблицы.
		/// </summary>
		private static readonly string SpreadSheetName = Settings.Default.SpreadSheetName;

		private readonly SheetsService _sheetsService;

		private readonly List<Sheet> _sheets = new List<Sheet>();

		public string SpreadSheetId { get; private set; }

		public CGoogleSpreadSheet(string spreadSheetId, string clientId, string clientSecret)
		{
			SpreadSheetId = spreadSheetId;
			var userCredential = GetSheetCredentials(clientId, clientSecret);
			_sheetsService = GetSheetsService(userCredential);
		}

		#region CredentialsAndService

		private static UserCredential GetSheetCredentials(string clientId, string clientSecret)
		{
			try
			{
				var credentialPath = Path.Combine(Directory.GetCurrentDirectory(), TokenFile);
				return GoogleWebAuthorizationBroker.AuthorizeAsync(
					new ClientSecrets
					{
						ClientId = clientId,
						ClientSecret = clientSecret
					}, 
					ScopesSheets,
					DefaultUser,
					CancellationToken.None,
					new FileDataStore(credentialPath, true)).Result;
			}
			catch (Exception ex)
			{
				var message = $"{nameof(GetSheetCredentials)}: {ex.Message}";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				return null;
			}
		}

		private static SheetsService GetSheetsService(UserCredential userCredential)
		{
			if (userCredential == null)
			{
				Program.ConsoleLog($"{nameof(GetSheetsService)}: userCredential is null");
				return null;
			}
			
			return new SheetsService(new BaseClientService.Initializer
			{
				HttpClientInitializer = userCredential,
				ApplicationName = ApplicationName
			});
		}

		#endregion

		/// <summary>
		/// Проверяет наличие всех необходимых листов в таблице.
		/// </summary>
		public void UpdateTableSheets(List<CPostgreServer> postgreServers)
		{
			var allSheets = CGoogleSheet.GetAllSheets(_sheetsService, SpreadSheetId);
			var needUpdateSheets = CGoogleSheet.CheckAndCreateSheets(_sheetsService, SpreadSheetId, allSheets, postgreServers);

			if (needUpdateSheets)
			{
				allSheets = CGoogleSheet.GetAllSheets(_sheetsService, SpreadSheetId);
			}

			_sheets.Clear();

			if (allSheets?.Any() ?? false)
			{
				_sheets.AddRange(allSheets);
			}
		}

		/// <summary>
		/// Проверяет актуальность SpreadSheetId, создает новую таблицу в случае неактуального SpreadSheetId, указанного в конфигурационном файле.
		/// </summary>
		public void ActualizeSpreadSheetId()
		{
			try
			{
				_sheetsService?.Spreadsheets.Get(SpreadSheetId).Execute();
			}
			catch
			{
				var message = $"{nameof(ActualizeSpreadSheetId)}: {nameof(SpreadSheetId)} = {SpreadSheetId} is invalid";
				Program.ConsoleLog(message);
				Program.Logger.Error(message);
				SpreadSheetId = CreateNewSpreadSheet(_sheetsService);
				if (SpreadSheetId == null)
				{
					return;
				}

				var newSpreadSheetMessage = $"{nameof(ActualizeSpreadSheetId)}: new spreadSheetId = {SpreadSheetId}";
				Program.ConsoleLog(newSpreadSheetMessage);
				Program.Logger.Warn(newSpreadSheetMessage);
			}
		}

		/// <summary>
		/// Создает новую таблицу и возвращает ее id.
		/// </summary>
		private static string CreateNewSpreadSheet(SheetsService sheetsService)
		{
			var spreadSheet = new Spreadsheet
			{
				Properties = new SpreadsheetProperties
				{
					Title = SpreadSheetName
				}
			};

			var request = sheetsService.Spreadsheets.Create(spreadSheet);
			request.Fields = "spreadsheetId";

			try
			{
				var newSpreadSheet = request.Execute();
				return newSpreadSheet.SpreadsheetId;
			}
			catch (Exception e)
			{
				var message = $"{nameof(CreateNewSpreadSheet)}: {e.Message}";
				Program.ConsoleLog(message);
				Program.Logger.Error(message);
				return null;
			}
		}

		/// <summary>
		/// Обновляет таблицу.
		/// </summary>
		public void UpdateGoogleTable(IEnumerable<CPostgreServer> postgreServers)
		{
			foreach (var server in postgreServers)
			{
				var sheet = _sheets.FirstOrDefault(x => string.Equals(x.Properties.Title, server.ServerName));
				if (sheet == null)
				{
					continue;
				}

				var range = $"{server.ServerName}!A1:D";
				var rows = CGoogleSheet.GetSheetRows(_sheetsService, range, SpreadSheetId, out var isSuccess);

				if (!isSuccess)
				{
					continue;
				}

				var sheetId = sheet.Properties.SheetId;
				if (rows == null && !CGoogleRow.AddRow(_sheetsService, SpreadSheetId, sheetId, ERowType.HeaderRow, server, 0))
				{
					continue;
				}
				
				rows = CGoogleSheet.GetSheetRows(_sheetsService, range, SpreadSheetId, out _);
				if (rows == null)
				{
					continue;
				}

				var lastRowIndex = rows.Count == 1
					? 1
					: rows.Count - 1;
				var addRowResult = CGoogleRow.AddRow(_sheetsService, SpreadSheetId, sheetId, ERowType.DbSizeRow, server, lastRowIndex);
				lastRowIndex = addRowResult ? lastRowIndex + 1 : lastRowIndex;
				CGoogleRow.AddRow(_sheetsService, SpreadSheetId, sheetId, ERowType.FooterRow, server, lastRowIndex);
			}
		}
	}
}
