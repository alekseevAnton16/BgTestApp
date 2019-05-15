using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace BGTestApp
{
	public class CGoogleSheet
	{
		private const string ApplicationName = "BGTestApp";
		private static readonly string[] ScopesSheets = {SheetsService.Scope.Spreadsheets};

		private readonly UserCredential _userCredential;
		private readonly SheetsService _sheetsService;

		public string SpreadSheetId { get; }

		public string ClientSecretJsonFilePath { get; }

		public CGoogleSheet(string spreadSheetId, string clientSecretJsonFilePath)
		{
			SpreadSheetId = spreadSheetId;
			ClientSecretJsonFilePath = clientSecretJsonFilePath;
			_userCredential = GetSheetCredentials(clientSecretJsonFilePath);
			_sheetsService = GetSheetsService(_userCredential);
		}

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
			catch (Exception e)
			{
				CStatic.Logger.Error($"{nameof(GetSheetCredentials)}: {e.Message}");
				return null;
			}
		}

		private static SheetsService GetSheetsService(UserCredential userCredential)
		{
			if (userCredential == null)
			{
				return null;
			}
			
			return new SheetsService(new BaseClientService.Initializer
			{
				HttpClientInitializer = userCredential,
				ApplicationName = ApplicationName
			});
		}

		/// <summary>
		/// Получает все листы из таблицы.
		/// </summary>
		public IList<Sheet> GetAllSheets()
		{
			try
			{
				var spreadSheet = _sheetsService?.Spreadsheets.Get(SpreadSheetId).Execute();
				return spreadSheet?.Sheets;
			}
			catch (Exception e)
			{
				CStatic.Logger.Error($"{nameof(GetAllSheets)}: {e.Message}");
				return null;
			}
		}

		/// <summary>
		/// Создает лист в таблице.
		/// </summary>
		private void CreateSheet(string sheetName)
		{
			var addSheetRequest = new AddSheetRequest {Properties = new SheetProperties {Title = sheetName}};
			BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
			{
				Requests = new List<Request> {new Request {AddSheet = addSheetRequest}}
			};

			var batchUpdateRequest =
				_sheetsService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, SpreadSheetId);

			batchUpdateRequest.Execute();
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
