using System;
using System.Collections.Generic;
using System.Globalization;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace BGTestApp
{
	public static class CGoogleRow
	{
		private static readonly List<string> Headers = new List<string> {"Сервер", "База данных", "Размер в ГБ", "Дата обновления"};

		/// <summary>
		/// Создает строку
		/// </summary>
		public static bool CreateRow(SheetsService sheetsService, string spreadSheetId, int? sheetId, ERowType rowType, CPostgreServer server, int lastRowIndex)
		{
			switch (rowType)
			{
				case ERowType.HeaderRow:
					return CreateHeaderRow(sheetsService, spreadSheetId, sheetId);
				case ERowType.DbSizeRow:
					return AddDbSizeRow(sheetsService, spreadSheetId, sheetId, server, lastRowIndex);
				case ERowType.FooterRow:
					return CreateFooterRow(sheetsService, spreadSheetId, sheetId, server, lastRowIndex);
				default:
					return false;
			}
		}

		/// <summary>
		/// Создает строку с заголовками.
		/// </summary>
		private static bool CreateHeaderRow(SheetsService sheetsService, string spreadSheetId, int? sheetId)
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

		/// <summary>
		/// Добавляет строку в лист.
		/// </summary>
		/// <returns>Индекс последней строки в листе.</returns>
		private static bool AddDbSizeRow(SheetsService sheetsService, string spreadSheetId, int? sheetId, CPostgreServer server, int lastRowIndex)
		{
			if (server == null)
			{
				return false;
			}

			var request = GetRequestForServerAdd(server, sheetId, lastRowIndex);
			var batchUpdate = new BatchUpdateSpreadsheetRequest
			{
				Requests = new List<Request> {request}
			};

			try
			{
				sheetsService.Spreadsheets.BatchUpdate(batchUpdate, spreadSheetId).Execute();
				return true;
			}
			catch (Exception ex)
			{
				var message = $"{nameof(AddDbSizeRow)}: {ex.Message}";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				return false;
			}
		}

		/// <summary>
		/// Получает запрос на добавление строки с описанием сервера.
		/// </summary>
		private static Request GetRequestForServerAdd(CPostgreServer server, int? sheetId, int lastRowIndex)
		{
			var serverNameValue = new CellData
			{
				UserEnteredValue = new ExtendedValue
				{
					StringValue = server.ServerName
				}
			};

			var dbNameValue = new CellData
			{
				UserEnteredValue = new ExtendedValue
				{
					StringValue = server.DatabaseName
				}
			};

			var dbSizeValue = new CellData
			{
				UserEnteredValue = new ExtendedValue
				{
					StringValue = double.IsNaN(server.DatabaseSize) ? "неизвестно" : Math.Round(server.DatabaseSize, 1).ToString(CultureInfo.InvariantCulture)
				}
			};

			var changeDateValue = new CellData
			{
				UserEnteredValue = new ExtendedValue
				{
					StringValue = DateTime.Now.ToString("dd.MM.yyyy")
				}
			};

			var values = new List<CellData>{serverNameValue, dbNameValue, dbSizeValue, changeDateValue};

			return new Request
			{
				UpdateCells = new UpdateCellsRequest
				{
					Start = new GridCoordinate
					{
						SheetId = sheetId,
						RowIndex = lastRowIndex,
						ColumnIndex = 0,
					},
					Rows = new List<RowData> {new RowData {Values = values}},
					Fields = "userEnteredValue"
				}
			};
		}

		/// <summary>
		/// Создает строку с заголовками.
		/// </summary>
		private static bool CreateFooterRow(SheetsService sheetsService, string spreadSheetId, int? sheetId, CPostgreServer server, int lastRowIndex)
		{
			if (server == null)
			{
				return false;
			}

			var request = GetRequestForAddFooterRow(server, sheetId, lastRowIndex);
			var batchUpdate = new BatchUpdateSpreadsheetRequest
			{
				Requests = new List<Request> {request}
			};

			try
			{
				sheetsService.Spreadsheets.BatchUpdate(batchUpdate, spreadSheetId).Execute();
				return true;
			}
			catch (Exception ex)
			{
				var message = $"{nameof(CreateFooterRow)}: {ex.Message}";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				return false;
			}
		}

		/// <summary>
		/// Получает запрос на добавление строки с описанием сервера.
		/// </summary>
		private static Request GetRequestForAddFooterRow(CPostgreServer server, int? sheetId, int lastRowIndex)
		{
			var serverNameValue = new CellData
			{
				UserEnteredValue = new ExtendedValue
				{
					StringValue = server.ServerName
				}
			};
			
			var freeStateValue = new CellData
			{
				UserEnteredValue = new ExtendedValue
				{
					StringValue = "свободно"
				}
			};

			var freeSize = new CellData
			{
				UserEnteredValue = new ExtendedValue
				{
					StringValue = double.IsNaN(server.DatabaseSize) || double.IsNaN(server.ServerSize) 
						? "неизвестно"
						: Math.Round(server.ServerSize - server.DatabaseSize, 1).ToString(CultureInfo.InvariantCulture)
				}
			};

			var changeDateValue = new CellData
			{
				UserEnteredValue = new ExtendedValue
				{
					StringValue = DateTime.Now.ToString("dd.MM.yyyy")
				}
			};

			var values = new List<CellData>{serverNameValue, freeStateValue, freeSize, changeDateValue};

			return new Request
			{
				UpdateCells = new UpdateCellsRequest
				{
					Start = new GridCoordinate
					{
						SheetId = sheetId,
						RowIndex = lastRowIndex,
						ColumnIndex = 0,
					},
					Rows = new List<RowData> {new RowData {Values = values}},
					Fields = "userEnteredValue"
				}
			};
		}
	}
}
