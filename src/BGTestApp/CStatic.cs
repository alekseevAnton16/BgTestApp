using NLog;

namespace BGTestApp
{
	public static class CStatic
	{
		public static ILogger Logger = LogManager.GetCurrentClassLogger();
	}
}
