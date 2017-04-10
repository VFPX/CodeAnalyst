* Program....: FOXREF.H
* Date.......: February 26, 2002
* Abstract...: Header file for Source Reference application
* Changes....:

#define TAB					CHR(9)

* possible actions for starting FoxRef (see foxrefstart.prg)
#define MODE_REFERENCES		0
#define MODE_GOTODEF		1
#define MODE_LOOKUP			2

* this is what ProjectFile property is set to if we're searching folders
* rather than a specific Project
#define PROJECT_GLOBAL		"REF_GLOBAL"


* extension to append to Project Name
* when naming the reference table 
* If Project is called VisGift.PJX, then 
* reference table is called VisGift_ref.dbf
#define RESULT_EXT			"_ref"

* RESULT_EXT is append to this to become global_ref.dbf
#define GLOBAL_TABLE		"global"

#define FILE_TABLE			"RefFile"
#define DEF_TABLE			"RefDef"
#define ADDIN_TABLE			"RefAddIn"

* default report filename
#define REPORT_FILE			"foxrefresults.frx"

#define FONT_DEFAULT		"Tahoma,8,N"

#define TABLEFIND_ALIAS		"TableFind"
#define WORD_DELIMITERS		" ,.&()=[]'" + ["] + CHR(9)
#define METHOD_DELIMITERS	" ,&()=[]'" + ["] + CHR(9)


#define DEFINITION_DELIMITERS	" ,()" + CHR(9)

#define MAX_TOKENS			100

* The following are invalid in an alias name so we replace with '_'
#define INVALID_ALIAS_CHARS	" .-@#$%^&*()+={}[]:;?/<>,\|~`'" + ["]

* max number of characters a code line can be in 
* a grid before we cut it off
#define MAX_LINE_LENGTH		254

* these are the Types defined in RefAddIn Table
#define ADDINTYPE_FINDFILE		"FINDFILE"
#define ADDINTYPE_IGNOREFILE	"IGNOREFILE"
#define ADDINTYPE_FINDWINDOW	"FINDWINDOW"
#define ADDINTYPE_REPORT		"REPORT"
#define ADDINTYPE_MATCH			"MATCH"
#define ADDINTYPE_WILDMATCH		"WILDMATCH"
#define ADDINTYPE_FILETYPE		"FILETYPE"


* default file types if none specified in history (Common Source)
#define FILETYPES_DEFAULT	  "*.scx;*.vcx;*.prg;*.frx;*.lbx;*.mnx;*.dbc;*.qpr;*.h"
#define FILETYPES_DEFINITIONS "*.scx;*.vcx;*.prg;*.h;*.dbc;*.frx;*.lbx;*.mpr;*.spr;*.qpr"

* default search engines and their class libraries
#define FILETYPE_CLASS_DEFAULT		"RefSearch"
#define FILETYPE_LIBRARY_DEFAULT	"FoxRefSearch.prg"



* references into aFileTypes array
#define FILETYPE_EXTENSION	1
#define FILETYPE_ENGINE		2


* for nSearchType parameter in SearchText() method
#define SEARCHTYPE_NORMAL		0
#define SEARCHTYPE_METHOD		1
#define SEARCHTYPE_EXPR			2
#define SEARCHTYPE_TEXT			3


#DEFINE EDENV_FILENAME      1
#DEFINE EDENV_LENGTH		2
#DEFINE EDENV_READONLY		12
#DEFINE EDENV_SELSTART		17
#DEFINE EDENV_SELEND		18

#define SCOPE_PROJECT		'P'
#define SCOPE_FOLDER		'D'
#define SCOPE_FILE			'F'
#define SCOPE_CLASS			'C'

* Resource File ID
#define RESOURCE_ID			"FOXREF"

* File Types, as used by FoxRef.FileAction
#define FILEACTION_NODEFINITIONS	'N'
#define FILEACTION_DEFINITIONS		'D'
#define FILEACTION_ERROR			'E'
#define FILEACTION_INACTIVE			'X'
#define FILEACTION_STOP				'S'


* Reference types as used by *_ref.reftype
#define REFTYPE_INIT			'I'
#define REFTYPE_SEARCH			'S'
#define REFTYPE_RESULT			'R'
#define REFTYPE_NOMATCH			'N'
#define REFTYPE_ERROR			'E'
#define REFTYPE_LOG				'L'
#define REFTYPE_INACTIVE		'X'

* Where we found the results - used by *_ref.findtype
#define FINDTYPE_CODE			'C'
#define FINDTYPE_TEXT			'T'
#define FINDTYPE_NAME			'N'
#define FINDTYPE_EXPR			'E'
#define FINDTYPE_PROPERTYNAME	'P'
#define FINDTYPE_PROPERTYVALUE	'='
#define FINDTYPE_WINDOW			'W'
#define FINDTYPE_OTHER			'X'

* Definition Types used in RefDef.dbf
#define DEFTYPE_NONE			' '
#define DEFTYPE_PARAMETER		'P'
#define DEFTYPE_LOCAL			'L'
#define DEFTYPE_PRIVATE			'V'
#define DEFTYPE_PUBLIC			'G'
#define DEFTYPE_PROCEDURE		'F'
#define DEFTYPE_CLASS			'C'
#define DEFTYPE_DEFINE			'#'
#define DEFTYPE_PROPERTY		'='
#define DEFTYPE_INCLUDEFILE		'I'
#define DEFTYPE_SETCLASSPROC	'S'


#define LOG_PREFIX	"*** "

* -- The following are used in evaluating report files (frx)

* Report FRX objtypes
#define RPTTYPE_HEADER		1
#define RPTTYPE_DBF			2
#define RPTTYPE_INDEX		3
#define RPTTYPE_RELATION	4
#define RPTTYPE_LABEL		5
#define RPTTYPE_LINE		6
#define RPTTYPE_BOX			7
#define RPTTYPE_GET			8
#define RPTTYPE_BAND		9
#define RPTTYPE_GROUP		10
#define RPTTYPE_PICTURE		17
#define RPTTYPE_VAR			18
#define RPTTYPE_FONT		23
#define RPTTYPE_DATAENV		25
#define RPTTYPE_DERECORD	26


* Report FRX band objcodes (objtype=9)
#define RPTCODE_TITLE		0
#define RPTCODE_PGHEAD		1
#define RPTCODE_COLHEAD		2
#define RPTCODE_GRPHEAD		3
#define RPTCODE_DETAIL		4
#define RPTCODE_GRPFOOT		5
#define RPTCODE_COLFOOT		6
#define RPTCODE_PGFOOT		7
#define RPTCODE_SUMMARY		8

#define MATCH_ANY			'?'
#define MATCH_MORE			'*'

#define EXPORTTYPE_DBF			"DBF"
#define EXPORTTYPE_TXT			"TXT"
#define EXPORTTYPE_HTML			"HTML"
#define EXPORTTYPE_XML			"XML"
#define EXPORTTYPE_XLS			"XLS"
#define EXPORTTYPE_CLIPBOARD	"CLIPBOARD"

* default extension for backups
#define BACKUP_EXTENSION	"bak"

* used by the ShellTo method
#define SW_HIDE             0
#define SW_SHOWNORMAL       1
#define SW_NORMAL           1
#define SW_SHOWMINIMIZED    2
#define SW_SHOWMAXIMIZED    3
#define SW_MAXIMIZE         3
#define SW_SHOWNOACTIVATE   4
#define SW_SHOW             5
#define SW_MINIMIZE         6
#define SW_SHOWMINNOACTIVE  7
#define SW_SHOWNA           8
#define SW_RESTORE          9
#define SW_SHOWDEFAULT      10
#define SW_FORCEMINIMIZE    11
#define SW_MAX              11
#define SE_ERR_NOASSOC 		31

#define HELPTOPIC_GENERAL	"Code References Window"
#define HELPTOPIC_REGEXPR	"Regular Expressions and Operators"


* -- BEGIN LOCALIZATION STRINGS ---

* Name of this application -- used in the caption of MessageBox() functions
#define APPNAME_LOC				"Code References"


* Used by Progress form
#define PROGRESS_SEARCHING_LOC	 "Searching"
#define PROGRESS_REFRESHING_LOC	 "Refreshing"
#define PROGRESS_DEFINITIONS_LOC "Collecting Definitions - "

#define SEARCH_CANCEL_LOC		"Are you sure you want to cancel the search?"
#define CANCEL_LOC				"Cancel"
#define ERROR_LOC				"ERROR"

* these are used in the Left Hand Pane of the results window (don't localize #SYMBOL#)
#define ALLRESULTS_LOC			"All Results"
#define REPLACELOG_LOC			"Replace Log: #SYMBOL#"
#define ALLLOGS_LOC				"Replacement Logs"
#define EMPTYTEXT_LOC			"(nothing)"

#define SCOPE_PROJECT_LOC		"Project"
#define SCOPE_FOLDER_LOC		"Folder"
#define SCOPE_FILE_LOC			"Current File"
#define SCOPE_CLASS_LOC			"Current Class"

* Search Comments and non comments (0), Exclude Comments (1), or Comments Only (2)
#define COMMENTS_INCLUDE	0
#define COMMENTS_EXCLUDE	1
#define COMMENTS_ONLY		2		

* used on the toolbar buttons
#define TOOLBAR_FIND_LOC		"Find"
#define	TOOLBAR_REFRESH_LOC		"Refresh"
#define TOOLBAR_OPTIONS_LOC		"Options"
#define TOOLBAR_REPLACE_LOC		"Replace"


#define SLOW_WARNING_LOC		"The search string you specified could result in abnormally long search times and" + CHR(10) + "consume large amounts of disk space."
#define WILDCARD_WARNING_LOC	"You cannot specify a search pattern consisting only of wildcards."
#define CONTINUE_LOC			"Would you like to continue?"

* Results Window caption
#define RESULTSTITLE_LOC		"Code References"
#define FOLDERNOEXIST_LOC		"The specified folder does not exist."

#define NAME_LOC				"Name"
#define EXPRESSION_LOC			"Expression"
#define COMMENT_LOC				"Comment"
#define PEMNAME_LOC				"Property/Method Name"

* Descriptions of what we might find in a report
#define OBJECTTYPE_HEADER_LOC	"Header"
#define OBJECTTYPE_DBF_LOC		"DBF"
#define OBJECTTYPE_INDEX_LOC	"Index"
#define OBJECTTYPE_RELATION_LOC	"Relation"
#define OBJECTTYPE_LABEL_LOC	"Label"
#define OBJECTTYPE_LINE_LOC		"Line"
#define OBJECTTYPE_BOX_LOC		"Box"
#define OBJECTTYPE_GET_LOC		"Get"
#define OBJECTTYPE_TITLE_LOC	"Title"
#define OBJECTTYPE_PGHEAD_LOC	"Page Header"
#define OBJECTTYPE_COLHEAD_LOC	"Column Header"
#define OBJECTTYPE_GRPHEAD_LOC	"Group Header"
#define OBJECTTYPE_DETAIL_LOC	"Detail Band"
#define OBJECTTYPE_GRPFOOT_LOC	"Group Footer"
#define OBJECTTYPE_COLFOOT_LOC	"Column Footer"
#define OBJECTTYPE_PGFOOT_LOC	"Page Footer"
#define OBJECTTYPE_SUMMARY_LOC	"Summary"
#define OBJECTTYPE_BAND_LOC		"Band"
#define OBJECTTYPE_GROUP_LOC	"Group"
#define OBJECTTYPE_PICTURE_LOC	"Picture"
#define OBJECTTYPE_VAR_LOC		"Variable"
#define OBJECTTYPE_FONT_LOC		"Font"
#define OBJECTTYPE_DATAENV_LOC	"Data Environment"
#define OBJECTTYPE_DERECORD_LOC	"Data Environment Record"
#define OBJECTTYPE_UNKNOWN_LOC	"Unknown"

#define FRX_ONENTRYEXPR_LOC		"On Entry Expression"
#define FRX_ONEXITEXPR_LOC		"On Exit Expression"
#define FRX_PRINTONLYWHEN_LOC	"Print Only When Expression"
#define FRX_VARSTORE_LOC		"Value to Store"
#define FRX_VARINITIAL_LOC		"Initial Value"

* Descriptions of what we might find in a table (dbf)
#define DBF_FIELDNAME_LOC			"Field Name"
#define DBF_FIELDVALIDATIONEXPR_LOC	"Field Validation Expression"
#define DBF_FIELDVALIDATIONTEXT_LOC	"Field Validation Text"
#define DBF_DEFAULTVALUE_LOC		"Default Value"
#define DBF_TABLEVALIDATIONEXPR_LOC	"Record Validation Rule"
#define DBF_TABLEVALIDATIONTEXT_LOC	"Record Validation Text"
#define DBF_TABLELONGNAME_LOC		"Long Table Name"
#define DBF_INSERTTRIGGER_LOC		"Insert Trigger Expression"
#define DBF_UPDATETRIGGER_LOC		"Update Trigger Expression"
#define DBF_DELETETRIGGER_LOC		"Delete Trigger Expression"
#define DBF_TAGNAME_LOC				"Index Tag Name"
#define DBF_TAGEXPR_LOC				"Index Tag Expression"
#define DBF_TAGFILTER_LOC			"Index Tag Filter"


* Descriptions of what we might find in a database (dbc)
#define DBC_STOREDPROCEDURE_LOC		"Stored Procedure"
#define DBC_COMMENT_LOC				"Database Comment"

#define DBC_VIEWNAME_LOC			"View Name"
#define DBC_VIEWSQL_LOC				"View SQL Code"
#define DBC_VIEWCOMMENT_LOC			"View Comment"
#define DBC_VIEWPARAMETERS_LOC		"View Parameters"
#define DBC_VIEWCONNECTNAME_LOC		"Connection Name"
#define DBC_VIEWRULEEXPR_LOC		"View Rule Expression"
#define DBC_VIEWRULETEXT_LOC		"View Rule Text"

#define DBC_CONNECTNAME_LOC			"Connection Name"
#define DBC_CONNECTSTRING_LOC		"Connection String"
#define DBC_CONNECTCOMMENT_LOC		"Connection Comment"
#define DBC_CONNECTDATABASE_LOC		"Connection Database"
#define DBC_CONNECTDATASOURCE_LOC	"Connection Data Source"
#define DBC_CONNECTUSERID_LOC		"Connection User ID"
#define DBC_CONNECTPASSWORD_LOC		"Connection Password"


* Descriptions of what we might find in a menu (mnx)
#define MNX_NAME_LOC				"Name"
#define MNX_PROMPT_LOC				"Prompt"
#define MNX_COMMAND_LOC				"Command"
#define MNX_MESSAGE_LOC				"Message"
#define MNX_PROCEDURE_LOC			"Procedure"
#define MNX_SETUP_LOC				"Setup"
#define MNX_CLEANUP_LOC				"Cleanup"
#define MNX_SKIPFOR_LOC				"Skip For"
#define MNX_RESOURCE_LOC			"Resource"
#define MNX_COMMENT_LOC				"Comment"

* Grid headers on the results form
#define GRID_FILENAME_LOC			"File name"
#define GRID_CLASSMETHOD_LOC		"Class.Method, Line"
#define GRID_CODE_LOC				"Code"

#define REPLACE_LOC					"Replace"
#define REPLACE_NOCHECKS_LOC  		"You must first select items in which you want to perform replacements."
#define REPLACE_REFRESH_LOC			"Do you want to refresh the results?"
#define REPLACE_NOTSUPPORTED1_LOC	"Some of the replacements selected are not supported by Code References" + CHR(10) + "(data structure changes, property names & values)."
#define REPLACE_NOTSUPPORTED2_LOC	"The activity log which follows the code replacement operation" + CHR(10) + "includes instructions on making these changes."
#define REPLACE_SKIPPED				"SKIPPED"
#define REPLACE_NOTSUPPORTED_LOC	"REPLACEMENT NOT SUPPORTED"
#define REPLACE_READONLY_LOC		"The file is read-only."

* when user selects "Go To Definition" and no definition exists, display this message
#define NODEFINITION_LOC			"Definition not found for the following:"
#define GOTODEFINITION_LOC			"Go To Definition"


* Descriptions of the Export Types
#define EXPORT_DBF_LOC				"Visual FoxPro Table (DBF)"
#define EXPORT_TXT_LOC				"Comma Delimited Text (TXT)"
#define EXPORT_HTML_LOC				"HyperText Markup Language (HTML)"
#define EXPORT_XML_LOC				"Extensible Markup Language (XML)"
#define EXPORT_XLS_LOC				"Microsoft Excel"
#define EXPORT_CLIPBOARD_LOC		"Clipboard"

* valid values for XMLExportType
#define	XMLFORMAT_ELEMENTS			1
#define	XMLFORMAT_ATTRIBUTES		2


* used by FileMatch method in FoxRefEngine
#define FILEMATCH_ANY		'?'
#define FILEMATCH_MORE		'*'


* What to display if there are no matching records
* when printing results report
#define PRINT_NOTHING_LOC			"There are no results to report on."

* What we display in the Search Comments combo box on the Find dialog
#define COMMENTS_INCLUDE_LOC		"Include Comments"
#define COMMENTS_EXCLUDE_LOC		"Ignore Comments"
#define COMMENTS_ONLY_LOC			"Comments Only"

#define CRITERIA_MATCHCASE_LOC		"Match Case"
#define CRITERIA_WHOLEWORDS_LOC		"Whole Words Only"
#define CRITERIA_FORMPROPERTIES_LOC	"Search Property Names/Values"
#define CRITERIA_PROJECTHOMEDIR_LOC	"Project Home Directory Only"
#define CRITERIA_SUBFOLDERS_LOC		"Search Subfolders"
#define CRITERIA_WILDCARDS_LOC		"Use Regular Expressions"

* used in FoxRefResults as titles for our criteria and filetypes
#define SEARCHOPTIONS_LOC			"Options"
#define FOLDER_LOC					"Folder"
#define PROJECT_LOC					"Project"
#define ERRORHEADER_LOC				"Warnings/Errors:"
#define SUMMARY_NOMATCHES_LOC		"No matches found"
#define REPLACEMENTTEXT_LOC			"Replacement Text"
#define DATETIME_LOC				"Date/Time"

* don't localize <MATCHCNT>, <FILECNT>, or <FILENAME> -- they are placeholders!
#define SUMMARY1_LOC				"<MATCHCNT> matches found in <FILECNT> files"
#define SUMMARY2_LOC				"<MATCHCNT> match found in <FILECNT> files"
#define SUMMARY3_LOC				"<MATCHCNT> matches found in <FILECNT> file"
#define SUMMARY4_LOC				"<MATCHCNT> match found in <FILECNT> file"
#define SUMMARY5_LOC				"<MATCHCNT> match found in <FILENAME>"
#define SUMMARY6_LOC				"<MATCHCNT> matches found in <FILENAME>"

* errors that could occur when creating or opening *_Ref table
#define ERROR_BADREFTABLE_LOC		"The Reference Table found to use is owned by another application:"
#define ERROR_CREATEREFTABLE_LOC	"The Reference Table could not be created"
#define ERROR_OPENREFTABLE_LOC		"The Reference Table could not be opened"

#define ERROR_BADDEFTABLE_LOC		"The Definition Table found to use is owned by another application:"
#define ERROR_CREATEDEFTABLE_LOC	"The Definition Table could not be created:"
#define ERROR_OPENDEFTABLE_LOC		"The Definition Table could not be opened:"

#define ERROR_BADFILETABLE_LOC		"The File Table found to use is owned by another application:"
#define ERROR_CREATEFILETABLE_LOC	"The File Table could not be created:"
#define ERROR_OPENFILETABLE_LOC		"The File Table could not be opened:"

#define ERROR_REPLACE_LOC			"Error encountered during replacement"
#define ERROR_NOBACKUP_LOC			"Backup could not be created so replacement not performed"
#define ERROR_FILENOTFOUND_LOC		"File not found"
#define ERROR_FILEMODIFIED_LOC		"Unable to update because file has been modified"
#define ERROR_WRITE_LOC				"Error encountered saving file (may be read-only)"
#define ERROR_NOENGINE_LOC			"Search/replace engine not defined for this file type"
#define ERROR_PATTERN_LOC			"The search engine does not support the specified pattern."
#define ERROR_SEARCHENGINE_LOC		"An error occurred performing the search." + CHR(10) + "The search pattern may be invalid."

* errors that can occur
#define ERROR_OPENFILE_LOC			"Unable to open file (file may be in use)"
#define ERROR_NOTFORM_LOC			"Not a form or class library file"
#define ERROR_NOTMENU_LOC			"Not a menu file"
#define ERROR_NOTREPORT_LOC			"Not a report or label file"
#define ERROR_EXCLUSIVE_LOC			"Unable to open tables exclusive for Cleanup."
#define ERROR_FOXTOOLS_LOC			"Unable to locate FOXTOOLS.FLL library."

* confirmation when Cleanup button is pressed on Options dialog
#define CLEANUP_CONFIRM_LOC			"Are you sure you want to remove all unused file references" + CHR(10) + "and definitions from the Code References tables?"

* used in FoxRefPrint when error encountered
#define ERROR_PRINT_LOC				"Error encountered running report."

* when "Clear All Results" is selected, confirm with this prompt
#define CLEARALL_LOC				"Are you sure you want to clear all result sets and replacement logs?"
#define CLEARALL_CAPTION_LOC		"Clear All"
#define CLEARALLRESULTS_LOC			"Are you sure you want to clear all result sets?"
#define CLEARALLRESULTS_CAPTION_LOC	"Clear All Results"

* used in FoxRefReplaceConfirm.scx to show what we're replacing
#define REPLACECONFIRM_REPLACE_LOC	"Replace:"
#define REPLACECONFIRM_WITH_LOC		"with:"

#define BACKUP_PREFIX_LOC			"Backup of"

* right-click menu options from FoxRefResults.scx -> ShowRightClickMenu()
#define MENU_DESCRIPTIONS_LOC		"\<Display Descriptions"
#define MENU_ALWAYSONTOP_LOC		"\<Always on Top"
#define MENU_OPENPROJECT_LOC		"\<Open Project"
#define MENU_SEARCH_LOC				"\<Search"
#define MENU_GLOBALREPLACE_LOC		"\<Replace"
#define MENU_REFRESH_LOC			"Re\<fresh"
#define MENU_PRINT_LOC				"\<Print"
#define MENU_EXPORT_LOC				"\<Export"
#define MENU_OPTIONS_LOC			"Opt\<ions"

#define MENU_COPY_LOC				"\<Copy"
#define MENU_CLEAR_LOC				"Clea\<r Result"
#define MENU_CLEARALL_LOC			"Clear \<All Results"

#define MENU_OPEN_LOC				"\<Open"
#define MENU_GOTODEFINITION_LOC		"\<View Definition"
#define MENU_SELECTALL_LOC			"Select \<All"
#define MENU_DESELECTALLWIN_LOC		"Clear All \<Displayed Selections"
#define MENU_DESELECTALL_LOC		"C\<lear All Selections"
#define MENU_SORT_LOC				"Sor\<t by"

#define MENUSORT_FILENAME_LOC		"File \<Name"
#define MENUSORT_CHECKED_LOC		"\<Selected"
#define MENUSORT_CLASSMETHOD_LOC	"\<Class.Method"
#define MENUSORT_METHOD_LOC			"\<Method"
#define MENUSORT_FILETYPE_LOC		"File \<Type"
#define MENUSORT_LOCATION_LOC		"\<Location"

#define MENU_EXPANDALL_LOC			"\<Expand All"
#define MENU_COLLAPSEALL_LOC		"C\<ollapse All"
#define MENU_SORTMOSTRECENT_LOC		"Sort by Most Recent First"

* used in place of Filename during Goto Definition when
* the definitions actually came from the open window
#define OPENWINDOW_LOC				"Open Window"

#define FILEEXISTS_LOC				"#FILENAME# already exists." + CHR(10) + "Do you want to replace it?"
#define FILE_INVALID_LOC			"The file name specified is invalid."
#define SAVEAS_LOC					"Save As"

#define FILETYPE_ALL_LOC			"All Files (*.*)"
#define FILETYPE_COMMON_LOC			"Common (*.scx;*.vcx;*.prg;*.frx;*.lbx;*.mnx;*.dbc;*.qpr;*.h)"
#define FILETYPE_SOURCE_LOC			"All Source (*.scx;*.vcx;*.prg;*.frx;*.lbx;*.mnx;*.dbc;*.dbf;*.cdx;*.qpr;*.h)"
#define FILETYPE_FORMSCLASSES_LOC	"Forms and Classes (*.scx;*.vcx;*.prg)"
#define FILETYPE_REPORTS_LOC		"Reports and Labels (*.frx;*.lbx)"
#define FILETYPE_MENUS_LOC			"Menus (*.mnx)"
#define FILETYPE_PROGRAMS_LOC		"Programs (*.prg;*.h;*.qpr;*.mpr)"
#define FILETYPE_DATA_LOC			"Data (*.dbc;*.dbf;*.cdx)"
#define FILETYPE_TEXT_LOC			"Text (*.txt;*.xml;*.xsl;*.htm;*.html;*.log;*.asp;*.aspx)"

* used in the Regular Expression help menu on the Search dialog
#define REGEXPR_SINGLECHAR_LOC		". Any single character"
#define REGEXPR_ZEROORMORE_LOC		"* Zero or more"
#define REGEXPR_ONEORMORE_LOC		"+ One or more"
#define REGEXPR_BEGINNINGOFLINE_LOC	"^ Beginning of line"
#define REGEXPR_ENDOFLINE_LOC		"$ End of line"
#define REGEXPR_ANYONECHAR_LOC		"[] Any one character in the set"
#define REGEXPR_NOTANYONECHAR_LOC	"[^] Any one character not in the set"
#define REGEXPR_OR_LOC				"| OR"
#define REGEXPR_ESCAPE_LOC			"\\ Escape Special Character"
#define REGEXPR_HELP_LOC			"Help on Regular Expressions"

* Used in FoxRefOptions dialog
#define FOXREFDIRECTORY_NOEXIST_LOC	"The specified folder for code references tables does not exist." + CHR(10) + CHR(10) + "Are you sure this is correct?"

#define WSH_REQUIRED_LOC			"The Windows Script Host must be installed for regular expression searches."

#define NODE_LOADING_LOC			"(loading...)"

#define CLASS_HEADER_LOC			"Class"