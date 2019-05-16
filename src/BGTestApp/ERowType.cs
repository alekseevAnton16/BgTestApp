namespace BGTestApp
{
	/// <summary>
	/// Тип строки.
	/// </summary>
	public enum ERowType
	{
		/// <summary>
		/// Заголовок.
		/// </summary>
		HeaderRow,

		/// <summary>
		/// Строка с описанием размера базы данных.
		/// </summary>
		DbSizeRow,

		/// <summary>
		/// Строка с описанием свободного места на с сервере.
		/// </summary>
		FooterRow,
	}
}
