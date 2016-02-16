<%

var util = require('util');

var TYPES = {
	TinyInt: function() {return {id: 16, type: TYPES.TinyInt}},
	SmallInt: function() {return {id: 2, type: TYPES.SmallInt}},
	Int: function() {return {id: 3, type: TYPES.Int}},
	BigInt: function() {return {id: 20, type: TYPES.BigInt}},
	VarChar: function(length) {return {id: 200, type: TYPES.VarChar, length: length}},
	Bit: function() {return {id: 11, type: TYPES.Bit}},
	DateTime: function() {return {id: 135, type: TYPES.DateTime}},
	Text: function() {return {id: 201, type: TYPES.Text}},
	NText: function() {return {id: 203, type: TYPES.NText}},
	Decimal: function(precision, scale) {return {id: 14, type: TYPES.Decimal, precision: precision, scale: scale}},
	Numeric: function(precision, scale) {return {id: 131, type: TYPES.Numeric, precision: precision, scale: scale}},
	Double: function() {return {id: 5, type: TYPES.Double}},
	NVarChar: function(length) {return {id: 202, type: TYPES.NVarChar, length: length}},
	Binary: function() {return {id: 128, type: TYPES.Binary}},
	Image: function() {return {id: 205, type: TYPES.Image}},
	VarBinary: function(length) {return {id: 204, type: TYPES.VarBinary, length: length}}
};

for (var type in TYPES) {
	(function(type) {
		TYPES[type].inspect = function inspect() {return '[sql.'+ type +']'};
	})(type);
}

getTypeByValue = function getTypeByValue(value) {
	if (value == null) return TYPES.NVarChar;

	switch (typeof value) {
		case 'string': return TYPES.NVarChar;
		case 'number': return TYPES.Int;
		case 'boolean': return TYPES.Bit;
		case 'object':
			if (Date.isDate(value)) return TYPES.DateTime;
			if (Buffer.isBuffer(value)) return TYPES.VarBinary;
			return TYPES.NVarChar;
		default: return TYPES.NVarChar;
	};
};

valueCorrection = function valueCorrection(value, metadata) {
	if (metadata.type === 135) {
		value = new Date(value);
	} else if (metadata.type === 128 || metadata.type === 204 || metadata.type === 205) {
		value = new Buffer(value);
	};
	return value;
};

Connection = function Connection(config, done) {
	this._config = config;
	
	this._connection = Server.CreateObject("ADODB.Connection");
	this._connection.ConnectionTimeout = (config.connectionTimeout || 15000) / 1000;
	this._connection.CommandTimeout = (config.requestTimeout || 240000) / 1000;
	this._connection.CursorLocation = 3;
	this._connection.Provider = 'SQLOLEDB';
	
	if ('function' === typeof done) return this.connect(done);
};

Connection.define({
	file: '/aspjs_modules/mssql/mssql.asp'
}, {
	connect: function connect(done) {
		var error = null;
		
		try {
			this._connection.Open('Provider=SQLOLEDB;Data Source='+ this._config.server +','+ (this._config.port || 1433) +';Initial Catalog='+ this._config.database +';Network Library=DBMSSOCN;', this._config.user, this._config.password);
		} catch (ex) {
			ex.stack = ex.stack || Error.captureStackTrace();
			error = ex;
		};
		
		util.defer(done, error);
		
		return this;
	},
	close: function close(done) {
		var error = null;
		
		try {
			this._connection.Close();
		} catch (ex) {
			ex.stack = ex.stack || Error.captureStackTrace();
			error = ex;
		};
		
		util.defer(done, error);

		return this;
	}
});

Request = function Request(connection) {
	this._command = Server.CreateObject("ADODB.Command");
	this._command.Parameters.Append(this._command.CreateParameter('@RETURN_VALUE', 3, 4, 0, null));
	this._command.ActiveConnection = connection._connection;

	this.parameters = {};
};

Request.define({
	file: '/aspjs_modules/mssql/mssql.asp'
}, {
	input: function input(name, type, value) {
		try {
			if (arguments.length === 2) {
				value = type;
				type = getTypeByValue(value);
			};
			
			if (type instanceof Function) type = type();
			if (type === TYPES.DateTime && value instanceof Date) value = value.toISOString();
			if (value === undefined) value = null;
			if (value !== value) value = null;
			
			if (type.id === 200 || type.id === 201 || type.id === 202 || type.id === 203 || type.id === 204) {
				value = String(value);
				type.length = type.length || value.length;
				
				if (value.length > 4000 && type.id === 202 || type.id === 203) {
					var param = this._command.CreateParameter('@'+ name, type.id, 1, -1, null);
					this._command.Parameters.Append(param); param.AppendChunk(value);
				} else if (value.length > 8000) {
					var param = this._command.CreateParameter('@'+ name, type.id, 1, -1, null);
					this._command.Parameters.Append(param); param.AppendChunk(value);
				} else {
					this._command.Parameters.Append(this._command.CreateParameter('@'+ name, type.id, 1, type.length || 0, value));
				};
			} else {
				this._command.Parameters.Append(this._command.CreateParameter('@'+ name, type.id, 1, type.length || 0, value));
			};
			
			//console.log('mssql.input', name, type, value);
			this.parameters[name] = {
				name: name,
				type: type,
				io: 1,
				value: value
			};
		} catch (ex) {
			ex = new Error('Invalid input parameter \''+ name +'\'. '+ ex.message);
			ex.stack = Error.captureStackTrace();
			throw ex;
		};
		return this;
	},
	output: function input(name, type) {
		try {
			if (type instanceof Function) type = type();

			this._command.Parameters.Append(this._command.CreateParameter('@'+ name, type.id, 2, type.length || 0, null));
			
			//console.log('mssql.output', name, type);
			this.parameters[name] = {
				name: name,
				type: type,
				io: 2,
				value: null
			};
		} catch (ex) {
			ex = new Error('Invalid input parameter \''+ name +'\'. '+ ex.message);
			ex.stack = Error.captureStackTrace();
			throw ex;
		};
		return this;
	},
	execute: function execute(procedure, done) {
		this._command.CommandType = 4;
		this._command.CommandText = procedure;
		
		var rst = Server.CreateObject("ADODB.Recordset");
		rst.CacheSize = 50;
		rst.CursorLocation = 3;
		
		var error = null, recordset = null, returnValue = 0;
		
		try {
			rst.open(this._command);
		} catch (ex) {
			ex.stack = ex.stack || Error.captureStackTrace();
			error = ex;
		};
		
		returnValue = this._command.Parameters("@RETURN_VALUE").Value;

		if (!error) {
			recordset = [];
			
			if (rst.State === 1 && !rst.EOF) {
				while (!rst.EOF) {
					var row = {};
					for (var i = 0, l = rst.Fields.Count, f; i < l; i++) {
						f = rst.Fields(i);
						row[f.name] = valueCorrection(f.value, {type: f.type, precision: f.precision, scale: f.numericScale});
					};
					recordset.push(row);
					rst.moveNext();
				};
				
				rst.close();
			};
		};
		
		util.defer(done, error, recordset, returnValue);
		
		return this;
	},
	query: function query(command, done) {
		this._command.CommandType = 1;
		this._command.CommandText = command;
		
		var rst = Server.CreateObject("ADODB.Recordset");
		rst.CursorType = 0;
		rst.LockType = 1;
		
		var error = null, recordset = null;
		
		try {
			rst.open(this._command);
		} catch (ex) {
			ex.stack = ex.stack || Error.captureStackTrace();
			error = ex;
		};

		if (!error) {
			recordset = [];
			
			if (rst.State === 1 && !rst.EOF) {
				while (!rst.EOF) {
					var row = {};
					for (var i = 0, l = rst.Fields.Count, f; i < l; i++) {
						f = rst.Fields(i);
						console.log('x', f.name, f.type);
						row[f.name] = valueCorrection(f.value, {type: f.type, precision: f.precision, scale: f.numericScale});
					};
					recordset.push(row);
					rst.moveNext();
				};
				
				rst.close();
			};
		};
		
		util.defer(done, error, recordset);

		return this;
	}
});

module.exports = {
	Connection: Connection,
	Request: Request
};

for (var type in TYPES) {
	module.exports[type] = TYPES[type];
};

%>
