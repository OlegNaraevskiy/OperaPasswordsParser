/*===================================================================
 * Copyright (c) 2022 Oleg Naraevskiy                   Date: 09.2022
 * Version IDE: MS VS 2022
 * Designed by: Oleg Naraevskiy / noa.oleg96@gmail.com      [09.2022]
 *===================================================================*/

using CsvHelper.Configuration;

namespace OperaPasswordsParser.Classes
{
	/// <summary>
	/// Данные аутентификации
	/// </summary>
	public class PasswordInfo
	{
		/// <summary>
		/// Наименование
		/// </summary>
		public string Name { get; set; }

		/// <summary>
		/// Ссылка на конечный сайт
		/// </summary>
		public string URL { get; set; }

		/// <summary>
		/// Логин
		/// </summary>
		public string UserName { get; set; }

		/// <summary>
		/// Пароль
		/// </summary>
		public string Password { get; set; }

		public PasswordInfo(string name, string uRL, string userName, string password)
		{
			Name = name;
			URL = uRL;
			UserName = userName;
			Password = password;
		}

		public PasswordInfo()
		{
		}
	}

	/// <summary>
	/// Mapping for export to Excel shell
	/// </summary>
	public sealed class PasswordInfoMap : ClassMap<PasswordInfo>
	{
		public PasswordInfoMap()
		{
			Map(m => m.Name).Name("name");
			Map(m => m.URL).Name("url");
			Map(m => m.UserName).Name("username");
			Map(m => m.Password).Name("password");
		}
	}
}
