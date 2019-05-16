using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using BGTestApp.Properties;
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
		private const string TokenFile = "token.json";
		private const string DefaultUser = "user";

		private static readonly string[] ScopesSheets = {SheetsService.Scope.Spreadsheets};

		private readonly SheetsService _sheetsService;

		private readonly List<Sheet> _sheets = new List<Sheet>();

		public string SpreadSheetId { get; }
		
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

				var sheetId = sheet.Properties.SheetId;
				if (rows == null && !CGoogleRow.CreateRow(_sheetsService, SpreadSheetId, sheetId, ERowType.HeaderRow, server, 0))
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
				var addRowResult = CGoogleRow.CreateRow(_sheetsService, SpreadSheetId, sheetId, ERowType.DbSizeRow, server, lastRowIndex);
				lastRowIndex = addRowResult ? lastRowIndex + 1 : lastRowIndex;
				CGoogleRow.CreateRow(_sheetsService, SpreadSheetId, sheetId, ERowType.FooterRow, server, lastRowIndex);
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
