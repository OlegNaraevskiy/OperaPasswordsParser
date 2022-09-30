/*===================================================================
 * Copyright (c) 2022 Oleg Naraevskiy                   Date: 09.2022
 * Version IDE: MS VS 2022
 * Designed by: Oleg Naraevskiy / noa.oleg96@gmail.com      [09.2022]
 *===================================================================*/

using Aspose.Cells;
using CsvHelper;
using OperaPasswordsParser.Classes;
using OperaPasswordsParser.Classes.Enums;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OperaPasswordsParser
{
	internal class Program
	{
		static void Main(string[] args)
		{
			bool isExit = false;

			while (!isExit)
			{
				Console.WriteLine("=======================================");

				Console.Write("Введите путь к файлу CSV: ");
				string filePath = Console.ReadLine();
				var resultRead = CsvFileParser.Read(filePath);

				Console.WriteLine("=======================================");

				Console.WriteLine(resultRead.Message);

				if (resultRead.Code == (int)StatusCodeEnum.Success)
				{
					Console.Write("Введите название файла Excel для сохранения: ");
					string fileName = Console.ReadLine();

					Console.Write("Введите путь для сохранения: ");
					string fileDirectory = Console.ReadLine();

					var resultWrite = ExportDataToExcel.Export(fileName, fileDirectory, resultRead.PasswordInfos);

					Console.WriteLine(resultWrite.Message);
				}

				Console.WriteLine("=======================================");

				Console.Write("Для выхода из программы нажмите Q, чтобы повторить нажмите любую другую клавишу: ");
				var isExitButton = Console.ReadKey();
				Console.WriteLine();

				if (isExitButton.Key == ConsoleKey.Q)
				{
					isExit = true;
				}
				else
				{
					Console.Clear();
				}
			}
		}

		/// <summary>
		/// Чтение CSV файла
		/// </summary>
		public static class CsvFileParser
		{
			/// <summary>
			/// Чтение файла
			/// </summary>
			/// <param name="filePath">Путь к файлу</param>
			/// <returns>
			/// Результат:	Успех - Список значений из CSV файла
			///				Ошибка - Описание ошибки
			/// </returns>
			public static ResultMessage Read(string filePath)
			{
				filePath = filePath.Trim();
				ResultMessage resultMessage = new ResultMessage();

				try
				{
					if (!File.Exists(filePath))
					{
						resultMessage.Code = (int)StatusCodeEnum.ReadError;
						filePath = string.IsNullOrWhiteSpace(filePath) ? "Адрес файла пуст" : filePath;
						resultMessage.Message = $"Исходного файла по пути: { filePath  } - не найден либо неверный формат пути.";

						return resultMessage;
					}

					List<PasswordInfo> resultList = new List<PasswordInfo>();

					using (var reader = new StreamReader(filePath))
					{
						using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
						{
							csv.Context.RegisterClassMap<PasswordInfoMap>();

							resultList = csv.GetRecords<PasswordInfo>().ToList();

							if (resultList != null && resultList.Count > 0)
							{
								resultMessage.Code = (int)StatusCodeEnum.Success;
								resultMessage.Message = StatusCodeEnum.Success.ToString();
								resultMessage.PasswordInfos = resultList;
							}
							else
							{
								resultMessage.Code = (int)StatusCodeEnum.ReadError;
								resultMessage.Message = StatusCodeEnum.ReadError.ToString();
								resultMessage.PasswordInfos = null;
							}
						}
					}
				}
				catch (Exception ex)
				{
					resultMessage.ExceptionInfo = ex;
					resultMessage.Message = ex.Message;
				}

				return resultMessage;
			}
		}

		#region Another way to parse (warning: no errors processing)
		/*public static class ParseWithTextFieldParser
		{
			public static void TryParse(string filePath = null)
			{
				using (TextFieldParser parser = new TextFieldParser(filePath))
				{
					parser.TextFieldType = FieldType.Delimited;
					parser.SetDelimiters(",");
					while (!parser.EndOfData)
					{
						string[] fields = parser.ReadFields();
						foreach (string field in fields)
						{
							//Действия с записями
						}
					}
					parser.Close();
				}
			}
		}*/
		#endregion

		/// <summary>
		/// Сохранение данных из CSV файла в форматированном Excel
		/// </summary>
		public static class ExportDataToExcel
		{
			/// <summary>
			/// Экспорт данных в файл
			/// </summary>
			/// <param name="fileName">Наименование файла</param>
			/// <param name="fileDirectory">Путь к директории</param>
			/// <param name="passwordInfos">Исходные данные</param>
			/// <returns>
			/// Результат:	Успех - Файл в формате Excel
			///				Ошибка - Описание ошибки
			/// </returns>
			public static ResultMessage Export(string fileName, string fileDirectory, List<PasswordInfo> passwordInfos)
			{
				fileName = fileName.Trim();
				fileDirectory =	fileDirectory.Trim();
				fileDirectory = fileDirectory.TrimEnd('\\');

				string filePath = $"{ fileDirectory }\\{ fileName }.xlsx";

				ResultMessage resultMessage = new ResultMessage();

				try
				{
					if (File.Exists(filePath))
					{
						resultMessage.Code = (int)StatusCodeEnum.WriteError;
						resultMessage.Message = $"Файл с таким наименованием уже существует.";

						return resultMessage;
					}
					else
					{
						using (File.Create(filePath));
					}

					using (Workbook workbook = new Workbook(filePath))
					{
						//Настройка формата таблицы
						ImportTableOptions tableOptions = new ImportTableOptions
						{
							CheckMergedCells = true,
							IsFieldNameShown = true
						};

						int rowCounts = workbook.Worksheets[0].Cells.ImportCustomObjects(passwordInfos, 1, 0, tableOptions);

						//+1 т.к. заголовок считывается из свойств класса
						if (rowCounts == passwordInfos.Count + 1)
						{
							workbook.Save($"{filePath}", SaveFormat.Xlsx);

							resultMessage.Code = (int)StatusCodeEnum.Success;
							resultMessage.Message = $"Файл успешно сохранён по пути: {filePath}";
						}
						else
						{
							resultMessage.Code = (int)StatusCodeEnum.WriteError;
							resultMessage.Message = $"Файл НЕ сохранён по пути: {filePath}. Расхождение по количеству строк.";
						}
					}
					
				}
				catch (Exception ex)
				{
					resultMessage.ExceptionInfo = ex;
					resultMessage.Message = ex.Message;
				}

				return resultMessage;
			}
		}
	}
}
