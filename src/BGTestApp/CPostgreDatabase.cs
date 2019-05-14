using System;
using Newtonsoft.Json.Linq;
using Npgsql;

namespace BGTestApp
{
	public class CPostgreDatabase
	{
		private const string GetSizeFunctionName = "select pg_size_pretty(pg_database_size";

		private NpgsqlConnection _connection;

		public string ServerName { get; private set; }

		public string ConnectionString { get; private set; }

		public string DbSize { get; private set; }

		public void UpdateDbSize()
		{
			try
			{
				_connection.Open();
				var command = new NpgsqlCommand($"{GetSizeFunctionName}('{_connection.Database}'))", _connection);
				var result = command.ExecuteScalar();
				_connection.Close();
				DbSize = (string) result;
			}
			catch (Exception ex)
			{
				CStatic.Logger.Error($"{nameof(UpdateDbSize)}: {ex.Message}");
				DbSize = null;
			}
		}

		/// <summary>
		/// Получает экземпляр <see cref="CPostgreDatabase"/> из json.
		/// </summary>
		public static CPostgreDatabase GetConnectionFromJson(string json)
		{
			try
			{
				return GetConnectionFromJson(JObject.Parse(json));
			}
			catch (Exception exception)
			{
				CStatic.Logger.Error($"{nameof(GetConnectionFromJson)}: {exception.Message}");
				return null;
			}
		}

		public static CPostgreDatabase GetConnectionFromJson(JObject json)
		{
			var host = json.ContainsKey("DatabaseHost") ? (string) json["DatabaseHost"] : null;
			var port = json.ContainsKey("DatabasePort") ? (int) json["DatabasePort"] : default(int);
			var database = json.ContainsKey("DatabaseName") ? (string) json["DatabaseName"] : null;
			var userId = json.ContainsKey("UserId") ? (string) json["UserId"] : null;
			var password = json.ContainsKey("Password") ? (string) json["Password"] : null;
			if (string.IsNullOrWhiteSpace(host) || port == default(int) || string.IsNullOrWhiteSpace(database) || string.IsNullOrWhiteSpace(userId)
			    || string.IsNullOrWhiteSpace(password))
			{
				return null;
			}

			var connectionString = $"Host={host};Port={port};Database={database};User Id={userId};Password={password};";

			var postgreDatabase = new CPostgreDatabase
			{
				ServerName = json.ContainsKey("ServerName") ? (string) json["ServerName"] : null,
				ConnectionString = connectionString,
				_connection = new NpgsqlConnection(connectionString)
			};
			postgreDatabase.UpdateDbSize();

			return postgreDatabase;
		}
	}
}
