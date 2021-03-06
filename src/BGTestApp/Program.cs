﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace BGTestApp
{
	public class Program
	{
		private const string ClientSecret = "client_secret.json";
		private static readonly string[] ScopesSheets = {SheetsService.Scope.Spreadsheets};
		private const string ApplicationName = "BGTestApp";
		private const string SpreadSheetId = "1KewdQFRR898O743PqMfOZTcpWouQT6kZVM0oGEvirlo";
		private const string Range = "'Server1'!A1:A";

		private static readonly string[,] Data =
		{
			{"11", "12", "13"},
			{"21", "22", "23"},
			{"31", "32", "33"},
		};

		public static void Main(string[] args)
		{
			var credential = GetSheetCredentials();
			var service = GetSheetsService(credential);
			UpdateSpreadSheet(service, SpreadSheetId, Data);
			Console.WriteLine("Complete");
			Console.ReadLine();
		}

		private static UserCredential GetSheetCredentials()
		{
			using (var fileStream = new FileStream(ClientSecret, FileMode.Open, FileAccess.Read))
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
