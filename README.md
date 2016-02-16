# asp.js SQL Server module

## Quick Example

```asp
<!--#INCLUDE VIRTUAL="/aspjs_modules/mssql/index.asp"-->
<%

var sql = require("mssql");

var conn = new sql.Connection({
	user: "...",
	password: "...",
	server: "localhost",
	database: "..."
}, function(err) {
	// ... error checks
	
	new sql.Request(conn)
	.input("param", "abc")
	.execute("test_sp", function(err, recordset, returnValue) {
		// ... error checks
		
		console.log(recordset);
	});
	
	new sql.Request(conn)
	.query("select newid() as newid", function(err, recordset) {
		// ... error checks
		
		console.log(recordset);
	});
});

%>
```

## Documentation

Inspired by [node-mssql](https://github.com/patriksimek/node-mssql).

<a name="license" />
## License

Copyright (c) 2016 Patrik Simek

The MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
