using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using BGTestApp.Properties;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json.Linq;

namespace BGTestApp
{
	public class Program
	{
		private const string ApplicationName = "BGTestApp";

		private static readonly string[] ScopesSheets = {SheetsService.Scope.Spreadsheets};
		private static readonly string SpreadSheetId = Settings.Default.SpreadSheetId;

		private static readonly string[,] Data =
		{
			{"11", "12", "13"},
			{"21", "22", "23"},
			{"31", "32", "33"},
		};

		public static void Main(string[] args)
		{
			var databases = GetDatabases();
			if (databases != null)
			{
				var credential = GetSheetCredentials();
				var service = GetSheetsService(credential);
				var allSheets = GetAllSheets(service);
				//CreateSheet(service, nameof());
				//UpdateSpreadSheet(service, SpreadSheetId, Data);
				Console.WriteLine("Complete");
				Console.ReadLine();
			}
		}
		
		/// <summary>
		/// Получает список <see cref="CPostgreDatabase"/> из json.
		/// </summary>
		private static List<CPostgreDatabase> GetDatabases()
		{
			JArray connectionStringsArray;
			try
			{
				connectionStringsArray = JArray.Parse(Settings.Default.ConnectionStrings);
			}
			catch (Exception ex)
			{
				CStatic.Logger.Error($"{nameof(GetDatabases)}: {ex.Message}");
				Console.WriteLine("Ошибка при чтении json из файла конфигурации");
				return null;
			}
			
			var result = new List<CPostgreDatabase>();
			foreach (var json in connectionStringsArray)
			{
				if (json is JObject jsonObject)
				{
					var postgreDb = CPostgreDatabase.GetConnectionFromJson(jsonObject);
					if (postgreDb != null)
					{
						result.Add(postgreDb);
					}
				}
			}

			return result;
		}

		private static UserCredential GetSheetCredentials()
		{
			using (var fileStream = new FileStream(Settings.Default.ClientSecretJsonFileName, FileMode.Open, FileAccess.Read))
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

		private static SheetsService GetSheetsService(UserCredential userCredential) => new SheetsService(new BaseClientService.Initializer
		{
			HttpClientInitializer = userCredential,
			ApplicationName = ApplicationName
		});

		private static IList<Sheet> GetAllSheets(SheetsService sheetsService)
		{
			var spreadSheet = sheetsService.Spreadsheets.Get(SpreadSheetId).Execute();
			return spreadSheet.Sheets;
		}

		/// <summary>
		/// Создание нового листа.
		/// </summary>
		private static void CreateSheet(SheetsService sheetsService, string sheetName)
		{
			var addSheetRequest = new AddSheetRequest {Properties = new SheetProperties {Title = sheetName}};
			BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
			{
				Requests = new List<Request> {new Request {AddSheet = addSheetRequest}}
			};

			var batchUpdateRequest =
				sheetsService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, SpreadSheetId);

			batchUpdateRequest.Execute();
		}

		private static void UpdateSpreadSheet(SheetsService sheetsService, string spreadSheetId, string[,] data)
		{
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
