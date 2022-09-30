/*===================================================================
 * Copyright (c) 2022 Oleg Naraevskiy                   Date: 09.2022
 * Version IDE: MS VS 2022
 * Designed by: Oleg Naraevskiy / noa.oleg96@gmail.com      [09.2022]
 *===================================================================*/

using OperaPasswordsParser.Classes.Enums;
using System;
using System.Collections.Generic;

namespace OperaPasswordsParser.Classes
{
	public class ResultMessage
	{
		public int Code { get; set; }

		public string Message { get; set; }

		public Exception ExceptionInfo { get; set; }

		public List<PasswordInfo> PasswordInfos { get; set; }

		public ResultMessage(int code, string message, Exception exceptionInfo, List<PasswordInfo> passwordInfos)
		{
			Code = code;
			Message = message;
			ExceptionInfo = exceptionInfo;
			PasswordInfos = passwordInfos;
		}

		public ResultMessage()
		{
			Code = (int)StatusCodeEnum.Unknown;
			Message = "unknown error";
			ExceptionInfo = null;
			PasswordInfos = null;
		}
	}
}
