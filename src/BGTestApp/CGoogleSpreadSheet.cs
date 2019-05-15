using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace BGTestApp
{
	public class CGoogleSpreadSheet
	{
		private const string ApplicationName = "BGTestApp";
		
		private static readonly string[] ScopesSheets = {SheetsService.Scope.Spreadsheets};

		private readonly SheetsService _sheetsService;

		private readonly List<Sheet> _sheets = new List<Sheet>();

		public string SpreadSheetId { get; }
		
		public CGoogleSpreadSheet(string spreadSheetId, string clientSecretJsonFilePath)
		{
			SpreadSheetId = spreadSheetId;
			var userCredential = GetSheetCredentials(clientSecretJsonFilePath);
			_sheetsService = GetSheetsService(userCredential);
		}

		#region CredentialsAndService

		private static UserCredential GetSheetCredentials(string clientSecretJsonFilePath)
		{
			try
			{
				using (var fileStream = new FileStream(clientSecretJsonFilePath, FileMode.Open, FileAccess.Read))
				{
					var credentialPath = Path.Combine(Directory.GetCurrentDirectory(), "sheetsCreds.json");
					return GoogleWebAuthorizationBroker.AuthorizeAsync(
						GoogleClientSecrets.Load(fileStream).Secrets,
						ScopesSheets,
						"user",
						CancellationToken.None,
						new FileDataStore(credentialPath, true)).Result;
				}
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
		public void CheckTableSheets(List<CPostgreServer> postgreServers)
		{
			var allSheets = CGoogleSheet.GetAllSheets(_sheetsService, SpreadSheetId);
			if (CGoogleSheet.CheckAndCreateSheets(_sheetsService, SpreadSheetId, allSheets, postgreServers))
			{
				allSheets = CGoogleSheet.GetAllSheets(_sheetsService, SpreadSheetId);
			}

			_sheets.Clear();
			_sheets.AddRange(allSheets);
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

				if (rows == null && !CGoogleRow.CreateHeaderRow(_sheetsService, SpreadSheetId, sheet.Properties.Index))
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
				var addRowResult = CGoogleRow.AddRow(_sheetsService, SpreadSheetId, sheet.Properties.Index, server, lastRowIndex);
				lastRowIndex = addRowResult ? lastRowIndex + 1 : lastRowIndex;
				CGoogleRow.CreateFooterRow(_sheetsService, SpreadSheetId, sheet.Properties.Index, server, lastRowIndex);
			}
		}
		
		private static void UpdateSpreadSheet(SheetsService sheetsService, string spreadSheetId, string[,] data)
		{
			//todo
			var requests = new List<Request>();
			for (var i = 0; i < data.GetLength(0); i++)
			{
				var values = new List<CellData>();
				for (var j = 0; j < data.GetLength(1); j++)
				{
					values.Add(new CellData
					{
						UserEnteredValue = new ExtendedValue
						{
							StringValue = data[i, j]
						}
					});
				}

				requests.Add(new Request
				{
					UpdateCells = new UpdateCellsRequest
					{
						Start = new GridCoordinate
						{
							SheetId = 0,
							RowIndex = i,
							ColumnIndex = 0,
						},
						Rows = new List<RowData> {new RowData {Values = values}},
						Fields = "userEnteredValue"
					}
				});
			}

			var busr = new BatchUpdateSpreadsheetRequest
			{
				Requests = requests
			};

			sheetsService.Spreadsheets.BatchUpdate(busr, spreadSheetId).Execute();
		}
	}
}
