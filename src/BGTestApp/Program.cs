﻿using System;
using System.Collections.Generic;
using System.Threading;
using BGTestApp.Properties;
using Newtonsoft.Json.Linq;

namespace BGTestApp
{
	public class Program
	{
		private static void Main()
		{
			var databases = GetDatabases();
			Start(databases);
			Console.ReadLine();
		}

		private static void Start(List<CPostgreServer> databases)
		{
			if (databases == null)
			{
				ConsoleLog($"{nameof(Start)}: databaseList is null");
				return;
			}

			var timeout = Settings.Default.TimeoutInSeconds * 1000;
			while (true)
			{
				databases.ForEach(x => x.UpdateDbSize());
				Thread.Sleep(timeout);
			}
		}

		/// <summary>
		/// Получает список <see cref="CPostgreServer"/> из json.
		/// </summary>
		private static List<CPostgreServer> GetDatabases()
		{
			JArray connectionStringsArray;
			try
			{
				connectionStringsArray = JArray.Parse(Settings.Default.ConnectionStrings);
			}
			catch (Exception ex)
			{
				var message = $"{nameof(GetDatabases)}: {ex.Message}";
				CStatic.Logger.Error(message);
				ConsoleLog(message);
				return null;
			}
			
			var result = new List<CPostgreServer>();
			foreach (var json in connectionStringsArray)
			{
				if (json is JObject jsonObject)
				{
					var postgreDb = CPostgreServer.GetConnectionFromJson(jsonObject);
					if (postgreDb != null)
					{
						result.Add(postgreDb);
					}
				}
			}

			return result;
		}

		public static void ConsoleLog(string message)
		{
			Console.WriteLine($"{DateTime.Now:dd.MM.yyyy HH.mm.ss} {message}");
		}
	}
}
