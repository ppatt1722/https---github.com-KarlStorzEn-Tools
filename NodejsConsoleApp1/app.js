
var gpdpFormatterOptions = {
	logfile: '',
	testToReqXlsxFilePath : '',
	scriptsFolderPath : '',
	gpdpOutputPath : ''
}

function gpdpFormatter(options) {

	// prepare constants for execution	
	var valueColumnName = 'Requirements ID';
	var keyColumnName = 'Test ID';

	var WINSTON = require('winston');
	WINSTON.remove(WINSTON.transports.Console);
	if ((typeof options.logfile) != 'undefined') {
		WINSTON.add(WINSTON.transports.File, { filename: logfile, level: 'debug' });
		WINSTON.add(WINSTON.transports.Console, { level: 'warn' });
	}
	else {
		WINSTON.add(WINSTON.transports.Console, { level: 'debug' });
	}
	
	if ((typeof options.isTest) != 'undefined' && options.isTest) {
		return {
			// add methods for testing here
			testToReqMapping	: ManyToManyLookup,
			sheetToCSV			: SheetToCSV, 
			ParseCommentBlocks	: ParseCommentBlocks,
			logger				: WINSTON
		}
	}
	else {
		return {
			// public methods and properties
			formatTestResults	: Main
		}
	}
	
	function Main() {
	}
	
	function ManyToManyLookup(workbookPath, options) {
		WINSTON.log('verbose', 'ManyToManyLookup: Entered.');
		if (typeof require !== 'undefined') XLSX = require('xlsx');
		var dict = [];
		var workbook = XLSX.readFile(workbookPath);
		var aintThatSomeSheet = workbook.SheetNames[0];
		var worksheet = workbook.Sheets[aintThatSomeSheet];
		// get cell range
		var range = XLSX.utils.decode_range(worksheet["!ref"]);
		var val, R, C, keyColumn = 0, valueColumn = 1, key, value, lastRow = 1;
		
		// iterate rows
		for (R = range.s.r; R <= range.e.r; ++R) {
			// iterate columns
			for (C = range.s.c; C <= range.e.c; ++C) {
				// get current cell value
				val = worksheet[XLSX.utils.encode_cell({ c: C, r: R })];
				// if header row get index of key and value columns
				if (R == 0) {
					if (val.v == options.keyColumnName) keyColumn = C;
					if (val.v == options.valueColumnName) valueColumn = C;
				} else { // not header row
					key = worksheet[XLSX.utils.encode_cell({ c: keyColumn, r: R })].v;
					value = worksheet[XLSX.utils.encode_cell({ c: valueColumn, r: R })].v;
					// initialize value with empty array if not already
					if (dict[key] == null)
						dict[key] = [];
					// push value as one-of-many mapped values for given key
					dict[key].push(value);
					break;
				}
			}
		}
		WINSTON.log('verbose', 'ManyToManyLookup: Exited.');
		return dict;
	}
	
	function SheetToCSV(workbookPath, sheet_name, options) {
		WINSTON.log('verbose', 'SheetToCSV: Entered.');
		var XLSX = require('xlsx');
		var workbook = XLSX.readFile(workbookPath);
		
		var aintThatSomeSheet = sheet_name;
		if (sheet_name == null)
			aintThatSomeSheet = workbook.SheetNames[0];
		
		/* Get worksheet */
		var worksheet = workbook.Sheets[aintThatSomeSheet];
		if (worksheet == null) console.log("The specified worksheet does not exist.");
		
		var csv = XLSX.utils.sheet_to_csv(worksheet);
		
		WINSTON.log('verbose', 'SheetToCSV: Exited.');
		return csv;
	}

	function ParseCommentBlocks(openingDelimiter, closingDelimiter, filename) {
		WINSTON.log('verbose', 'ParseCommentBlocks: Entered.');
		var fs = require('fs');
		var ENCODING = require('encoding');

		var output = [];
		var delims = []; delims.push(openingDelimiter); delims.push(closingDelimiter);
		var index1 = 0; index2 = 0; state = false;
		var value = '';

		
		// This line opens the file as a readable stream
		var readStream = fs.createReadStream(filename);
		
		//readStream.on('readable', function() {
			var chunk;
			chunk = fs.readFileSync(filename).toString();
			//while (null != (chunk = readStream.read())) {
			//	chunk = chunk.toString();
				for (var i = 0, len = chunk.length; i < len; i++) {
					var header = accumulator(chunk[i]);
					if (header == null)
						continue;
					output.push(ParseCommentBlockMetadata(header));
				}
			//}

			if (state) {
				var msg = 'ParseCommentBlocks: Unclosed comment block encountered in ' + filename + ', aborting.';
				WINSTON.log('error', msg);
				throw Error(msg);
			}			
			WINSTON.log('verbose', 'ParseCommentBlocks: Exited.');
		//});
		
		return output;
		
		// parsing helper one character at a time
	    function accumulator(char) {
			var savedState = state;
			if (delims[index1][index2] == char) {
				if (index2 + 1 < delims[index1].length)
					index2++;
				else { // we completed a delimiter
					index2 = 0;
					state = index1 == 0 ? true : false;
					index1 = index1 == 1 ? 0 : 1;
					if (!state)				// if closing delim return text
					{
						WINSTON.log('verbose', 'ParseCommentBlocks: Closing delimiter detected.');
						return value;
					}
					else					// if opening delim reset
					{
						WINSTON.log('verbose', 'ParseCommentBlocks: Opening delimiter detected.');
						value = '';
					}
				}
			}
			else if (state) {				// only if state is in between delimiters...	
				index2 = 0;			
				value += char;				// ...do we accumulate char to block being formed
			}
			return null;					// indicate no block completed
		}
	}

	function ParseCommentBlockMetadata(blockText) {
		WINSTON.log('verbose', 'ParseCommentBlockMetadata: Entered.');
		var metadata = new Object();
		var lines = blockText.split("\n");
		for (var line in lines) {
			line = lines[line];
			if (line.startsWith('-- ')) {
				if (line.indexOf(':', 0) != -1) {
					line = line.substr(3);
					var spans = line.split(':');
					metadata[spans[0].trim()] = spans[1].trim();
					WINSTON.log('verbose', 'ParseCommentBlockMetadata: parsed ' + spans[0].trim() + ' : ' + spans[1].trim());
				}
			}
		}
		
		WINSTON.log('verbose', 'ParseCommentBlockMetadata: Exited.');
		return metadata;
	}
}

function gpdpFormatterTester(objectUnderTest, testOptions) {
	objectUnderTest.logger.log('verbose', 'Testing ManyToManyLookup, see results in console.');
	var map = objectUnderTest.testToReqMapping(testOptions.mappingXlsxPath, testOptions);
	
	for (var i in map)
		for (var j in map[i])
			console.log(i + ' => ' + map[i][j]);
	
	objectUnderTest.logger.log('verbose', 'Testing ParseCommentBlocks, see results in console.');
	var metadatas = objectUnderTest.ParseCommentBlocks(testOptions.openingDelim,
																testOptions.closingDelim,
																testOptions.scriptPath);
	for (var metadata in metadatas)
		for (prop in metadatas[metadata])
			objectUnderTest.logger.log('verbose', prop + ': ' + metadatas[metadata][prop]);
}

// these options work with tested object and its tester
var testOptions = {
	isTest			: true,
	mappingXlsxPath	: 'C:\\Users\\pspattillo\\Documents\\Doc\\Stafford Project\\testtoreq.xlsx',
	keyColumnName	: 'Test ID',
	valueColumnName	: 'Requirements ID',
	scriptPath      :  'C:\\Users\\pspattillo\\Documents\\Doc\\Stafford Project\\Sample_Script\\Buttons_spec.js',
	openingDelim	: '/*',
	closingDelim    : '*/'
}


var objectUnderTest = gpdpFormatter(testOptions);
objectUnderTest.logger.log('verbose', 'Invoke unit test for gpdpFormatted.');
gpdpFormatterTester(objectUnderTest, testOptions);