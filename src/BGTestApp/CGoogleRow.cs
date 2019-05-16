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
		/// Добавляет строку.
		/// </summary>
		public static bool AddRow(SheetsService sheetsService, string spreadSheetId, int? sheetId, ERowType rowType, CPostgreServer server, int lastRowIndex)
		{
			Request request = null;
			switch (rowType)
			{
				case ERowType.HeaderRow:
					request = GetRequestForAddHeaderRow(sheetId);
					break;
				case ERowType.DbSizeRow when server != null && sheetId != null && lastRowIndex != 0:
					request = GetRequestForAddDbSizeRow(server, sheetId, lastRowIndex);
					break;
				case ERowType.FooterRow when server != null && sheetId != null && lastRowIndex != 0:
					request = GetRequestForAddFooterRow(server, sheetId, lastRowIndex);
					break;
			}

			if (request == null)
			{
				return false;
			}

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
				var message = $"{nameof(AddRow)} (rowType = {rowType}): {ex.Message}";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				return false;
			}
		}
		
		/// <summary>
		/// Получение запроса на добавление заголовка.
		/// </summary>
		private static Request GetRequestForAddHeaderRow(int? sheetId)
		{
			var cellDataList = new List<CellData>();
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
				cellDataList.Add(value);
			}

			return new Request
			{
				UpdateCells = new UpdateCellsRequest
				{
					Start = new GridCoordinate
					{
						SheetId = sheetId,
						ColumnIndex = 0,
					},
					Rows = new List<RowData> {new RowData {Values = cellDataList}},
					Fields = "userEnteredValue"
				}
			};
		}
		
		/// <summary>
		/// Получает запрос на добавление строки с описанием сервера.
		/// </summary>
		private static Request GetRequestForAddDbSizeRow(CPostgreServer server, int? sheetId, int lastRowIndex)
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
