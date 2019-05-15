using System;
using System.Collections.Generic;
using BGTestApp.Properties;
using Newtonsoft.Json.Linq;

namespace BGTestApp
{
	public class Program
	{
		public static void Main(string[] args)
		{
			var databases = GetDatabases();
			if (databases != null)
			{
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
	}
}
