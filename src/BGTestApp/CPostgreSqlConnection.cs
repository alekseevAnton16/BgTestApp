using System;
using Npgsql;

namespace BGTestApp
{
	public class CPostgreSqlConnection
	{
		private const string GetSizeFunctionName = "select pg_database_size";

		private static CPostgreSqlConnection _currentConnection;

		public string ConnectionString { get; }

		private readonly NpgsqlConnection _connection;
		
		public CPostgreSqlConnection(string connectionString)
		{
			ConnectionString = connectionString;
			_connection = new NpgsqlConnection(connectionString);
		}

		public void GetDbSize()
		{
			try
			{
				_connection.Open();
				var command = new NpgsqlCommand($"{GetSizeFunctionName}('{_connection.Database}')", _connection);
				var result = command.ExecuteScalar();
				_connection.Close();
			}
			catch (Exception ex)
			{
				//
			}
		}
	}
}
