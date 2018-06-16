﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Command15_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00000508-0000-0010-8000-00AA006D2EA4")]
	public interface Command15_Deprecated : _ADO
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Connection_Deprecated ActiveConnection { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		string CommandText { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		Int32 CommandTimeout { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		bool Prepared { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Parameters_Deprecated Parameters { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi.Enums.CommandTypeEnum CommandType { get; set; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		string Name { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="recordsAffected">object recordsAffected</param>
		/// <param name="parameters">optional object parameters</param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Recordset_Deprecated Execute(object recordsAffected, object parameters, object options);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="recordsAffected">object recordsAffected</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Recordset_Deprecated Execute(object recordsAffected);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="recordsAffected">object recordsAffected</param>
		/// <param name="parameters">optional object parameters</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Recordset_Deprecated Execute(object recordsAffected, object parameters);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		/// <param name="size">optional Int32 Size = 0</param>
		/// <param name="value">optional object value</param>
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Parameter_Deprecated CreateParameter(object name, object type, object direction, object size, object value);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Parameter_Deprecated CreateParameter();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Parameter_Deprecated CreateParameter(object name);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Parameter_Deprecated CreateParameter(object name, object type);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Parameter_Deprecated CreateParameter(object name, object type, object direction);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		/// <param name="size">optional Int32 Size = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		NetOffice.ADODBApi._Parameter_Deprecated CreateParameter(object name, object type, object direction, object size);

		#endregion
	}
}
