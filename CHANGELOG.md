## Version 2.3 (7/8/2020)

Changes in this release:

* CHANGE: Updated minimum .NET Framework to v4.5

## Version 2.2 (11/17/2011)

Changes in this release:

* FIX: Fix bug in ExportBamDefinitionXml that caused output XML to be truncated with larger BAM models (when the BAM XML was greater than ~32K in size).

NOTE: OLE DB export from XLSX files requires installation of Office 2007 Data Connectivity Components or Access Database Engine 2010 Redistributable.

This release also includes a command-line utility program called ExportBamDefinitionXml that can extract the BAM definition XML from an XLS/XLSX file and write it out to a file. This utility can be used to automatically keep a BAM XML in sync with a BAM XLS file during local or automated builds.

## Version 2.1 (5/3/2010)

Changes in this release:

* NEW: Support for Office Data Connectivity Components 2010
* NEW: Include both x86 and x64 EXE's due to lack of support in DCC 2010 for x86 and x64 side-by-side installation
* NEW: Add error return code to command-line app so that the caller can determine success vs. failure
* CHANGE: Add OLE DB provider fall-back process - DCC 2007, then DCC 2010, then Jet 4.0 (Jet 4.0 only supports XLS files)

NOTE: OLE DB export from XLSX files requires installation of Office 2007 Data Connectivity Components or Office 2010 Data Connectivity Components.

This release also includes a command-line utility program called ExportBamDefinitionXml that can extract the BAM definition XML from an XLS/XLSX file and write it out to a file. This utility can be used to automatically keep a BAM XML in sync with a BAM XLS file during local or automated builds.

## Version 2.0 (2/16/2010)

Changes in this release:

* NEW: Export functionality no longer requires Excel to be installed (uses OLE DB vs. Excel Automation; also enables usage inside a Windows service)
* NEW: Override the code generation entirely, if desired, by providing a path to a custom XSLT file on the command line
* NEW: Override the BAM database connection strings at runtime, if necessary, using a new activity class constructor (Direct, Buffered)
* NEW: Comprehensive documentation included
* CHANGE: Modified the generated code to set activity item values using indexer instead of Add() (enables multiple property sets of same item)
* CHANGE: Added explicit call to Flush() at the end of the CommitXActivity() method and a new public Flush() method on each activity class (Direct, Buffered)
* CHANGE: Optimized the generated code to minimize objects created at runtime
* CHANGE: Added partial and virtual to activity classes
* CHANGE: Changed container class (xESApi) to a static class
* CHANGE: Refreshed the generated code to use proper indentation, use XML comments on methods and use generic collection classes

NOTE: OLE DB export from XLSX files requires installation of Office 2007 Data Connectivity Components.

This release includes a command-line utility program called ExportBamDefinitionXml that can extract the BAM definition XML from an XLS/XLSX file and write it out to a file. This utility can be used to automatically keep a BAM XML in sync with a BAM XLS file during local or automated builds.

## Version 1.2.1 (8/7/2008)

This release includes the first checkin of the source code plus these changes and fixes:

* Support for a relative path to the Excel workbook vs. explicit path only
* Added the partial keyword to the generated API class
* Added a method to the activity classes to enable continuation and return a continuation ID
* Added a command-line parameter to specify the .NET namespace of the generated code
* Explicitly pass "en-us" culture info to the Excel methods to avoid a globalization bug reported by users
* Minor code improvements, removed XSLT resource class in favor of Visual Studio's built-in resource support
* Bug fix to AddCustomReference - the type and name parameters were reversed. Since people had probably noticed that and swapped their parameter values, the parameters have been reversed in the AddCustomReference interface and in the call to the EventStream. Double-check any AddCustomReference calls to be safe.

This release also includes a new command-line utility program called ExportBamDefinitionXml that can extract the BAM definition XML from an XLS file and write it out to a file. This utility can be used to automatically keep a BAM XML in sync with a BAM XLS file during local or automated builds.
