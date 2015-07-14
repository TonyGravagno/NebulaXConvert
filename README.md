##NebulaXConvert by Nebula Research and Development

Software/ReadMe v1.0 (July 13, 2015)

This utility converts source files to binary files for Microsoft Excel and other applications.

###Usage:
	NebulaXConvert.exe "FromFile.xml" "ToFile.xls"

Any file name is allowed. If the path is not specified the current folder is assumed. If the file or path includes a space, enclose the filename in quotes as seen above.

Source files can be **XML**, **CSV**, or another version of Excel.

Once in XLS or XLSX format, any application that supports these formats should be able to open the files, including online services and mobile apps.

If the target has a .xls extension, the result is a file compatible with Excel 2003.

If the target has a .xlsX extension, the file will be compatible with Excel 2007+.

If the source file is already in Excel 2003 format with a .XLS extension, it converts to Excel 2007+ XLSX.

If the source file is XLSX and the target is XLS, it converts to Excel 2003 format.

No other targets are supported at this time (CSV, PDF, XLSB(BIFF12), ODS, etc).

Microsoft Excel and the .NET Framework v4+ MUST be present on the system to perform the conversion.

The target format, 2007, 2010, 2013, Office 365, etc, is dependent on the version of Excel which is on the system running the utility. The XLSX target actually means "Convert to the latest version supported by the Excel installed on this system".

The output is simply 'yes' or 'no' based on the success of the conversion. This allows another utility to easily capture and process the result.

If the operation fails and the result is no, Error messages can be obtained by appending the letter 'e' as a final parameter to the command. (No dash required.) The output will then show:

	ERROR: Opening:
			or
	ERROR: Saving:
	
The next lines will be the error message. The output ends with:

	RESULT:
	no
			or
	RESULT:
	yes
	
The consistent output structure should allow for parsing if required.

Common errors include:
- The source or target file is already opened and Excel cannot operate on them
- The source file doesn't exist
- The extension is not XLS or XLSX
- The source file is invalid
- Excel is not installed
- .NET4 is not installed

Please report issues to the GitHub [issue tracker](https://github.com/TonyGravagno/NebulaXConvert/issues) and subscribe to the repo for updates.

If you like the software, a note would be appreciated @TonyGravagno.

This software is provided under the MIT license. See LICENSE.txt for terms.
