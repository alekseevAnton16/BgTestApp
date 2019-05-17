using System;
using Newtonsoft.Json.Linq;
using Npgsql;

namespace BGTestApp
{
	public class CPostgreServer
	{
		private const string GetSizeFunctionName = "select pg_database_size";

		private static readonly double _dbSizeToGbСoeff = Math.Pow(1024, 3);

		private NpgsqlConnection _connection;

		public string ServerName { get; private set; }

		public string DatabaseName { get; private set; }

		public string ConnectionString { get; private set; }

		public double DatabaseSize { get; private set; }

		public double ServerSize { get; private set; }

		/// <summary>
		/// Обновляет размер БД.
		/// </summary>
		public void UpdateDbSize()
		{
			try
			{
				_connection.Open();
				var command = new NpgsqlCommand($"{GetSizeFunctionName}('{_connection.Database}')", _connection);
				var result = command.ExecuteScalar();
				_connection.Close();
				DatabaseSize = result is long resultAsLong 
					? resultAsLong / _dbSizeToGbСoeff
					: double.NaN;
			}
			catch (Exception ex)
			{
				var message = $"{nameof(UpdateDbSize)}: {ex.Message}";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				DatabaseSize = double.NaN;
			}
		}
		
		/// <summary>
		/// Получает экземпляр <see cref="CPostgreServer"/> из json.
		/// </summary>
		public static CPostgreServer GetConnectionFromJson(JObject json)
		{
			try
			{
				var host = json.ContainsKey("ServerHost") ? (string) json["ServerHost"] : null;
				var port = json.ContainsKey("ServerPort") ? (int) json["ServerPort"] : default(int);
				var database = json.ContainsKey("DatabaseName") ? (string) json["DatabaseName"] : null;
				var userId = json.ContainsKey("UserId") ? (string) json["UserId"] : null;
				var password = json.ContainsKey("Password") ? (string) json["Password"] : null;
				if (string.IsNullOrWhiteSpace(host) || port == default(int) || string.IsNullOrWhiteSpace(database) || string.IsNullOrWhiteSpace(userId)
					|| string.IsNullOrWhiteSpace(password))
				{
					return null;
				}

				var connectionString = $"Host={host};Port={port};Database={database};User Id={userId};Password={password};";

				var postgreDatabase = new CPostgreServer
				{
					ServerName = json.ContainsKey("ServerName") ? (string) json["ServerName"] : null,
					DatabaseName = database,
					ConnectionString = connectionString,
					ServerSize = json.ContainsKey("ServerSize") ? (double) json["ServerSize"] : double.NaN,
					_connection = new NpgsqlConnection(connectionString)
				};
				postgreDatabase.UpdateDbSize();

				return postgreDatabase;
			}
			catch (Exception ex)
			{
				var message = $"{nameof(GetConnectionFromJson)}: {ex.Message}";
				Program.Logger.Error(message);
				Program.ConsoleLog(message);
				return null;
			}
		}
	}
}
