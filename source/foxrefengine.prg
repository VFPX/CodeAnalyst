* Abstract...:
*	Primary class for Code References application.
*
* Changes....:
*
#include "foxpro.h"
#include "foxref.h"

DEFINE CLASS FoxRef AS Session
	PROTECTED lIgnoreErrors AS Boolean
	PROTECTED lRefreshMode
	PROTECTED cProgressForm
	PROTECTED lCancel
	PROTECTED tTimeStamp
	PROTECTED lIgnoreErrors

	Name = "FoxRef"

	* search match engines
	MatchClass		   = "MatchDefault"
	MatchClassLib      = "foxmatch.prg"
	WildMatchClass     = "MatchWildcard"
	WildMatchClassLib  = "foxmatch.prg"

	* default search engine for Open Window
	FindWindowClass    = "RefSearchWindow"
	FindWindowClassLib = "FoxRefSearch_Window.prg"

	oSearchEngine   = .NULL.
	oWindowEngine   = .NULL.

	WindowHandle    = -1
	WindowFilename  = ''
	WindowLineNo    = 0
	

	Comments        = COMMENTS_INCLUDE
	MatchCase       = .F.
	WholeWordsOnly  = .F.
	ProjectHomeDir  = .F.  && True to search only files in Project's Home Directory or below

	SubFolders      = .F.
	Wildcards       = .F.
	Quiet           = .F.  && quiet mode -- don't display search progress
	ShowProgress    = .T.  && show a progress form

	Errors          = .NULL.

	FileTypes       = ''
	ReportFile      = REPORT_FILE
	
	XSLTemplate     = "foxref.xsl"

	* MRU array lists
	DIMENSION aLookForMRU[10]
	DIMENSION aReplaceMRU[10]
	DIMENSION aFolderMRU[10]
	DIMENSION aFileTypesMRU[10]
	DIMENSION aDefaultFileTypes[1]

	aLookForMRU     = ''
	aReplaceMRU     = ''
	aFolderMRU      = ''
	aFileTypesMRU   = ''

	Pattern             = ''
	OverwritePrior      = .T.
	ConfirmReplace      = .T.  && confirm each replacement
	BackupOnReplace     = .T.  && false to not backup when doing global replace
	DisplayReplaceLog   = .T.  && create activity log for replacements
	PreserveCase        = .F.  && preserve case during a Replace operation

	FoxRefDirectory = ''
	RefTable        = ''
	DefTable        = ''
	FileTable       = ''
	AddInTable      = ''
	ProjectFile     = ''
	FileDirectory   = ''
		
	ActivityLog     = ''
	
	* The following are set by the Options dialog
	* There should be a corresponding entry in FoxRefOption.DBF
	* (except for BackupStyle & FontString)
	IncludeDefTable     = .T.  && create Definition table when searching
	CodeOnly            = .F.  && search only source code & expressions (not names and other none-code items)
	FormProperties      = .T.  && search form/class property names & values
	AutoProjectHomeDir  = .F.  && True to search only files in Project's Home Directory or below when doing definitions automatically
	ShowRefsPerLine     = .F.  && display a column in search results that depicts number of references found on the line
	ShowFileTypeHistory	= .F.  && True to keep filetype history in addition to showing common filetypes in search dialog
	ShowDistinctMethodLine = .F. && True to show columns for method/line apart from Class
	SortMostRecentFirst = .F.

	BackupStyle         = 1    && 1 = "filename.ext.bak"   2 = "Backup of filename.ext"
	FontString          = FONT_DEFAULT
	
	* XML Export Options
	XMLFormat          = XMLFORMAT_ELEMENTS
	XMLSchema          = .T.

	* This is the SetID for the last Replacement Log after it's saved
	ReplaceLogSetID    = ''

	* properties used internally
	cSetID              = ''
	lRefreshMode        = .F.
	lIgnoreErrors       = .F.
	oProgressForm       = .NULL.
	lCancel             = .F.
	tTimeStamp          = .NULL.
	lDefinitionsOnly    = .F.

	* collection of files we've backed up in this session
	oBackupCollection   = .NULL.

	oFileCollection     = .NULL.
	oSearchCollection   = .NULL.
	oProcessedCollection = .NULL.

	oEngineCollection   = .NULL.
	oReportCollection   = .NULL.
	
	oFileTypeCollection = .NULL.
	
	cTalk           = ''
	nLangOpt        = 0
	cEscapeState    = ''
	cSYS3054        = ''
	cSaveUDFParms   = ''
	cSaveLib        = ''
	cExclusive      = ''
	cCompatible     = ''
	

	oOptions        = .NULL.
	lInitError      = .F.
	
	oProjectFileRef = .NULL.


	PROCEDURE Init(lRestorePrefs)
		LOCAL nSelect
		LOCAL oException
		LOCAL cAddInType
		LOCAL nMemoWidth
		
		THIS.cTalk = SET("TALK")
		SET TALK OFF
		SET DELETED ON
		

		THIS.cCompatible = SET("COMPATIBLE")		
		SET COMPATIBLE OFF
		
		THIS.cExclusive = SET("EXCLUSIVE")
		SET EXCLUSIVE OFF

		THIS.nLangOpt = _VFP.LanguageOptions
		_VFP.LanguageOptions = 0

		THIS.cEscapeState = SET("ESCAPE")
		SET ESCAPE OFF

		THIS.cSYS3054 = SYS(3054)
		SYS(3054,0)

		THIS.cSaveLib      = SET("LIBRARY")

		THIS.cSaveUDFParms = SET("UDFPARMS")
		SET UDFPARMS TO VALUE

		SET EXACT OFF

		* Changed on 02/14/2002 06:13:18 PM by Ryan - moved to FoxRefStart.prg
		* THIS.RestorePrefs()

		* collection of errors
		THIS.Errors = NEWOBJECT("CFoxRefCollection", "FoxRefCollection.prg")
		
		* create a collection that contains the files we've backup up
		THIS.oBackupCollection = NEWOBJECT("CFoxRefCollection", "FoxRefCollection.prg")
		THIS.oFileCollection   = NEWOBJECT("CFoxRefCollection", "FoxRefCollection.prg")
		THIS.oSearchCollection = NEWOBJECT("CFoxRefCollection", "FoxRefCollection.prg")
		THIS.oProcessedCollection = CREATEOBJECT("Collection")

		THIS.oEngineCollection = CREATEOBJECT("Collection")
		THIS.oReportCollection = CREATEOBJECT("Collection")
		THIS.oFileTypeCollection = CREATEOBJECT("Collection") && default filetypes from FoxRefAddin

		THIS.oOptions = NEWOBJECT("FoxResource", "FoxResource.prg")


		IF m.lRestorePrefs
			THIS.RestorePrefs()
		ENDIF
		
		THIS.FoxRefDirectory = THIS.FoxRefDirectory
		
	ENDFUNC


	PROCEDURE Destroy()
		LOCAL cCompatible

		THIS.CloseProgress()

		IF THIS.cEscapeState = "ON"
			SET ESCAPE ON		
		ENDIF
		IF THIS.cTalk = "ON"
			SET TALK ON	
		ENDIF
		IF THIS.cExclusive = "ON"
			SET EXCLUSIVE ON
		ENDIF
		SYS(3054,INT(VAL(THIS.cSYS3054)))

		_VFP.LanguageOptions = THIS.nLangOpt

		IF THIS.cSaveUDFParms = "REFERENCE"
			SET UDFPARMS TO REFERENCE
		ENDIF
		
		m.cCompatible = THIS.cCompatible
		SET COMPATIBLE &cCompatible
	ENDFUNC

	FUNCTION FoxRefDirectory_Assign(cFoxRefDirectory)
		m.cFoxRefDirectory = ADDBS(m.cFoxRefDirectory)
		IF EMPTY(m.cFoxRefDirectory) OR !DIRECTORY(m.cFoxRefDirectory)
			m.cFoxRefDirectory = ADDBS(HOME(7))
			IF !DIRECTORY(m.cFoxRefDirectory)
				m.cFoxRefDirectory = ADDBS(HOME())
			ENDIF
		ENDIF
		THIS.FoxRefDirectory = m.cFoxRefDirectory

		THIS.DefTable   = THIS.FoxRefDirectory + DEF_TABLE
		THIS.FileTable  = THIS.FoxRefDirectory + FILE_TABLE
		THIS.AddInTable = THIS.FoxRefDirectory + ADDIN_TABLE

		THIS.InitAddIns()
		THIS.OpenTables()
	ENDFUNC

	* Open the Add In table which contains our filetypes to process
	FUNCTION InitAddIns(lExclusive)
		LOCAL nSelect
		LOCAL nMemoWidth
		LOCAL cAddInType

		m.nSelect = SELECT()

		* if we don't find the AddIn table on disk, then copy out
		* our project version of it
		IF !FILE(FORCEEXT(THIS.AddInTable, "DBF"))
			TRY
				USE FoxRefAddin IN 0 SHARED AGAIN
				SELECT FoxRefAddIn
				COPY TO (THIS.AddInTable) WITH PRODUCTION
			CATCH
			FINALLY
				IF USED("FoxRefAddin")
					USE IN FoxRefAddIn
				ENDIF
			ENDTRY
		ENDIF
		
		TRY
			IF m.lExclusive
				USE (THIS.AddInTable) ALIAS AddInCursor IN 0 EXCLUSIVE
			ELSE
				USE (THIS.AddInTable) ALIAS AddInCursor IN 0 SHARED AGAIN
			ENDIF
		CATCH
		ENDTRY

		IF !USED("AddInCursor")
			* we didn't find the Add-In table on disk, so use our built-in version
			TRY
				USE FoxRefAddIn ALIAS AddInCursor IN 0 SHARED AGAIN
			CATCH
			ENDTRY
		ENDIF

		* process AddIn table
		THIS.oEngineCollection = CREATEOBJECT("Collection")
		THIS.oReportCollection = CREATEOBJECT("Collection")
		THIS.oFileTypeCollection = CREATEOBJECT("Collection") && default filetypes from FoxRefAddin

		THIS.AddFileType("*.*", FILETYPE_CLASS_DEFAULT, FILETYPE_LIBRARY_DEFAULT)

		IF USED("AddInCursor")
			m.nMemoWidth = SET("MEMOWIDTH")
			SET MEMOWIDTH TO 1000

			* setup file types
			SELECT AddInCursor
			SCAN ALL
				m.cAddInType = RTRIM(AddInCursor.Type)

				DO CASE
				CASE m.cAddInType == ADDINTYPE_FINDFILE
					THIS.AddFileType(AddInCursor.Data, AddInCursor.ClassName, AddInCursor.ClassLib)

				CASE m.cAddInType == ADDINTYPE_IGNOREFILE
					THIS.AddFileType(AddInCursor.Data)

				CASE m.cAddInType == ADDINTYPE_FINDWINDOW
					IF !EMPTY(AddInCursor.ClassName)
						THIS.FindWindowClass    = AddInCursor.ClassName
						THIS.FindWindowClassLib = AddInCursor.ClassLib
					ENDIF

				CASE m.cAddInType == ADDINTYPE_MATCH
					IF !EMPTY(AddInCursor.ClassName)
						THIS.MatchClass    = AddInCursor.ClassName
						THIS.MatchClassLib = AddInCursor.ClassLib
					ENDIF
						
				CASE m.cAddInType == ADDINTYPE_WILDMATCH
					IF !EMPTY(AddInCursor.ClassName)
						THIS.WildMatchClass    = AddInCursor.ClassName
						THIS.WildMatchClassLib = AddInCursor.ClassLib
					ENDIF

				CASE m.cAddInType == ADDINTYPE_REPORT
					THIS.AddReport(AddInCursor.Data, AddInCursor.ClassLib, AddInCursor.ClassName, AddInCursor.Method, AddInCursor.Filename)

				CASE m.cAddInType == ADDINTYPE_FILETYPE
					THIS.oFileTypeCollection.Add(AddInCursor.Data)
				ENDCASE
			ENDSCAN
			SET MEMOWIDTH TO (m.nMemoWidth)

		ENDIF
		SELECT (m.nSelect)			
	ENDFUNC

	PROCEDURE FontString_Access
		IF EMPTY(THIS.FontString)
			RETURN FONT_DEFAULT
		ELSE
			RETURN THIS.FontString
		ENDIF
	ENDPROC


	* Add another file to process definitions for
	* These are files that aren't in our normal
	* scope processing, but are discovered along the
	* way, such as #include
	FUNCTION AddFileToProcess(cFilename)
		IF FILE(m.cFilename)
			THIS.oFileCollection.AddNoDupe(LOWER(FULLPATH(m.cFilename)))
		ENDIF
	ENDFUNC
	
	* Add a file to search -- these are files that aren't
	* in our normal scope processing, but we search anyhow,
	* such as a Table within a DBC when searching a project
	FUNCTION AddFileToSearch(cFilename)
		IF FILE(m.cFilename)
			THIS.oSearchCollection.AddNoDupe(LOWER(FULLPATH(m.cFilename)))
		ENDIF
	ENDFUNC


	* Add a report from the Add-Ins
	FUNCTION AddReport(cReportName, cClassLib, cClassName, cMethod, cFilename)
		LOCAL oReportAddIn
		
		oReportAddIn = NEWOBJECT("ReportAddIn")
		oReportAddIn.ReportName      = m.cReportName
		oReportAddIn.RptClassLibrary = m.cClassLib
		oReportAddIn.RptClassName    = m.cClassName
		oReportAddIn.RptMethod       = m.cMethod
		oReportAddIn.RptFilename     = m.cFilename
		
		TRY
			THIS.oReportCollection.Add(oReportAddIn, m.cReportName)
		CATCH
		ENDTRY
	ENDFUNC

	* Add a filetype search
	FUNCTION AddFileType(cFileSkeleton, cClassName, cClassLibrary)
		LOCAL nIndex
		LOCAL lSuccess
		LOCAL oEngine

		IF VARTYPE(m.cClassName) <> 'C'
			m.cClassName = ''
		ENDIF

		m.cFileSkeleton = UPPER(ALLTRIM(m.cFileSkeleton))
		IF EMPTY(m.cFileSkeleton)
			m.cFileSkeleton = "*.*"
		ENDIF

		FOR m.i = 1 TO THIS.oEngineCollection.Count
			IF THIS.oEngineCollection.GetKey(m.i) == m.cFileSkeleton
				THIS.oEngineCollection.Remove(m.i)
				EXIT
			ENDIF
		ENDFOR

		IF EMPTY(m.cClassName)
			THIS.oEngineCollection.Add(.NULL., m.cFileSkeleton)
		ELSE
			TRY
				m.oEngine = NEWOBJECT(m.cClassName, m.cClassLibrary)
			CATCH
				m.oEngine = .NULL.
			FINALLY
			ENDTRY
			IF !ISNULL(oEngine)
				IF m.cFileSkeleton = "*.*" AND THIS.oEngineCollection.Count > 0
					THIS.oEngineCollection.Add(m.oEngine, m.cFileSkeleton, 1)
				ELSE
					THIS.oEngineCollection.Add(m.oEngine, m.cFileSkeleton)
				ENDIF
			ENDIF
		ENDIF
		
		RETURN m.lSuccess
	ENDFUNC

	* Initialize all of the search engines
	* with our search options
	PROCEDURE SearchInit()
		LOCAL i
		LOCAL oException
		LOCAL nSelect
		LOCAL lSuccess
		
		nSelect = SELECT()

		THIS.ClearErrors()

		m.oException = .NULL.
		m.lSuccess = .T.
		
		THIS.oProcessedCollection.Remove(-1)
		
		TRY
			IF THIS.Wildcards
				THIS.oSearchEngine = NEWOBJECT(THIS.WildMatchClass, THIS.WildMatchClassLib)
			ELSE
				THIS.oSearchEngine = NEWOBJECT(THIS.MatchClass, THIS.MatchClassLib)
			ENDIF
			m.lSuccess = THIS.oSearchEngine.InitEngine()
		CATCH TO oException
			m.lSuccess = .F.
			MessageBox(m.oException.Message, MB_ICONEXCLAMATION, APPNAME_LOC)
		ENDTRY


		IF m.lSuccess
			THIS.oSearchEngine.MatchCase      = THIS.MatchCase
			THIS.oSearchEngine.WholeWordsOnly = THIS.WholeWordsOnly

			IF THIS.oSearchEngine.SetPattern(THIS.Pattern)
				* IF THIS.WindowHandle >= 0
					* this create the engine for searching open windows
					TRY
						THIS.oWindowEngine = NEWOBJECT(THIS.FindWindowClass, THIS.FindWindowClassLib)
					CATCH TO oException
						MessageBox(m.oException.Message + CHR(10) + CHR(10) + THIS.FindWindowClass + " (" + THIS.FindWindowClassLib + ")", MB_ICONSTOP, APPNAME_LOC)
					ENDTRY
					IF VARTYPE(THIS.oWindowEngine) == 'O'
						WITH THIS.oWindowEngine
							.SetID           = THIS.cSetID
							.oSearchEngine   = THIS.oSearchEngine
							.Pattern         = THIS.Pattern
							.Comments        = THIS.Comments
							.IncludeDefTable = THIS.IncludeDefTable
							.FormProperties  = THIS.FormProperties
							.CodeOnly        = THIS.CodeOnly
							.PreserveCase    = THIS.PreserveCase
						ENDWITH
					ENDIF
				* ENDIF

				* collection of engines for various filetypes
				FOR m.i = 1 TO THIS.oEngineCollection.Count
					IF VARTYPE(THIS.oEngineCollection.Item(m.i)) == 'O'
						WITH THIS.oEngineCollection.Item(m.i)
							.SetID             = ''
							.oSearchEngine     = THIS.oSearchEngine
							.Pattern           = THIS.Pattern
							.Comments          = THIS.Comments
							.IncludeDefTable   = THIS.IncludeDefTable
							.FormProperties    = THIS.FormProperties
							.CodeOnly          = THIS.CodeOnly
							.PreserveCase      = THIS.PreserveCase
						ENDWITH
					ENDIF
				ENDFOR
				
			ELSE
				m.lSuccess = .F.
				MessageBox(ERROR_PATTERN_LOC, MB_ICONEXCLAMATION, APPNAME_LOC)
			ENDIF
		ENDIF
		
		RETURN m.lSuccess
	ENDPROC

	* Global Replace on all checked files in current RefTable
	FUNCTION GlobalReplace(cReplaceText)
		** Not required in code analyst
		RETURN 0
	ENDFUNC

	* take the original text, determine the case,
	* and apply it the new text
	FUNCTION SetCasePreservation(cOriginalText, cReplacementText)
		LOCAL i
		LOCAL ch
		LOCAL lLower
		LOCAL lUpper

		IF ISUPPER(LEFTC(m.cOriginalText, 1)) AND ISLOWER(SUBSTRC(m.cOriginalText, 2, 1))
			* proper case
			m.cReplacementText = PROPER(m.cReplacementText)
		ELSE
			m.lLower = .F.
			m.lUpper = .F.
			FOR m.i = 1 TO LENC(m.cOriginalText)
				ch = SUBSTRC(m.cOriginalText, m.i, 1)
				DO CASE
				CASE BETWEEN(ch, 'A', 'Z')
					m.lUpper = .T.
				CASE BETWEEN(ch, 'a', 'z')
					m.lLower = .T.
				ENDCASE
				
				IF m.lLower AND m.lUpper
					EXIT
				ENDIF
			ENDFOR
			
			DO CASE
			CASE m.lLower AND !m.lUpper
				m.cReplacementText = LOWER(m.cReplacementText)
			CASE !m.lLower AND m.lUpper
				m.cReplacementText = UPPER(m.cReplacementText)
			ENDCASE
		ENDIF
				
		RETURN m.cReplacementText
	ENDFUNC


	* Do a replacement on designated file
	FUNCTION ReplaceFile(cRefID, cReplaceText, lBackupOnReplace)
		LOCAL nSelect
		LOCAL oFoxRefRecord
		LOCAL lSuccess
		LOCAL oEngine
		LOCAL lBackup
		LOCAL cFilename
		LOCAL oException
		LOCAL oReplaceCollection
		
		m.lSuccess = .F.
		IF USED("FoxRefCursor") AND VARTYPE(cRefID) == 'C' 
			m.nSelect = SELECT()

			IF PCOUNT() < 3 OR VARTYPE(m.lBackupOnReplace) <> 'L'
				m.lBackupOnReplace = THIS.BackupOnReplace
			ENDIF

			IF SEEK(m.cRefID, "FoxRefCursor", "RefID")
				IF USED("FileCursor") AND SEEK(FoxRefCursor.FileID, "FileCursor", "UniqueID")
					m.cFilename = ADDBS(RTRIM(FileCursor.Folder)) + RTRIM(FileCursor.Filename)

					IF FILE(m.cFilename)
						m.oEngine = THIS.GetEngine(JUSTFNAME(m.cFilename))

						* if we don't have an engine object defined, then assume
						* we're supposed to ignore this filetype
						IF VARTYPE(m.oEngine) == 'O'
							WITH m.oEngine
								.PreserveCase = THIS.PreserveCase
								m.lSuccess = .T.

								* create a backup of the file if we haven't already
								* done so in this session
								IF m.lBackupOnReplace
									* m.cFilename = ADDBS(RTRIM(m.oFoxRefRecord.Folder)) + RTRIM(m.oFoxRefRecord.Filename)
									m.lBackup = .T.
									FOR m.i = 1 TO THIS.oBackupCollection.Count
										IF THIS.oBackupCollection.Item(m.i) == UPPER(m.cFilename)
											m.lBackup = .F.
											EXIT
										ENDIF
									ENDFOR
									IF m.lBackup
										IF .BackupFile(m.cFilename, THIS.BackupStyle)
											* add to collection so we don't backup again during this session
											THIS.oBackupCollection.Add(UPPER(m.cFilename))
										ELSE
											m.lSuccess = .F.
											THIS.AddError(m.cFilename + ": " + ERROR_NOBACKUP_LOC)
										ENDIF
									ENDIF
								ENDIF
								

								* do the replace
								IF m.lSuccess
									oReplaceCollection = CREATEOBJECT("Collection")

									SELECT FoxRefCursor
									SCAN ALL FOR RefID == cRefID
										SCATTER MEMO NAME oFoxRefRecord
										oReplaceCollection.Add(oFoxRefRecord)
									ENDSCAN
									m.oException = .ReplaceWith(m.cReplaceText, m.oReplaceCollection, m.cFilename)
									IF !ISNULL(m.oException)
										m.lSuccess = .F.
										THIS.AddError(m.cFilename + ": " + IIF(EMPTY(m.oException.UserValue), m.oException.Message, m.oException.UserValue))
										THIS.AddLog(LOG_PREFIX + ERROR_LOC + ": " + IIF(EMPTY(m.oException.UserValue), m.oException.Message, m.oException.UserValue))
									ENDIF
									THIS.AddLog(.ReplaceLog)
								ENDIF
							ENDWITH

						ELSE
							THIS.AddError(m.cFilename + ": " + ERROR_NOENGINE_LOC)
						ENDIF
					ELSE
						THIS.AddError(m.cFilename + ": " + ERROR_FILENOTFOUND_LOC)
					ENDIF
				ENDIF

			ENDIF			
			SELECT (m.nSelect)
		ENDIF

		

		RETURN m.lSuccess
	ENDFUNC




	* Methods for updating Errors collection	
	PROCEDURE ClearErrors()
		THIS.Errors.Remove(-1)
	ENDPROC

	PROCEDURE AddError(cErrorMsg)
		THIS.Errors.Add(cErrorMsg)
	ENDPROC

	PROCEDURE AddLog(cLog)
		IF PCOUNT() == 0
			IF !EMPTY(THIS.ActivityLog)
				THIS.ActivityLog = THIS.ActivityLog + CHR(13) + CHR(10)
			ENDIF
		ELSE
			IF !EMPTY(m.cLog)
				THIS.ActivityLog = THIS.ActivityLog + m.cLog + CHR(13) + CHR(10)
			ENDIF
		ENDIF
	ENDPROC

	
	* Save replacement log
	FUNCTION SaveLog(cReplaceText)
		IF VARTYPE(m.cReplaceText) <> 'C'
			m.cReplaceText = ''
		ENDIF
	
		THIS.ReplaceLogSetID = SYS(2015)
		INSERT INTO FoxRefCursor ( ;
		  UniqueID, ;
		  SetID, ;
		  RefID, ;
		  RefType, ;
		  FileID, ;
		  Symbol, ;
		  ClassName, ;
		  ProcName, ;
		  ProcLineNo, ;
		  LineNo, ;
		  ColPos, ;
		  MatchLen, ;
		  Abstract, ;
		  RecordID, ;
		  UpdField, ;
		  Checked, ;
		  NoReplace, ;
		  Timestamp, ;
		  Inactive ;
		 ) VALUES ( ;
		  SYS(2015), ;
		  THIS.ReplaceLogSetID, ;
		  SYS(2015), ;
		  REFTYPE_LOG, ;
		  '', ;
		  m.cReplaceText, ;
		  '', ;
		  '', ;
		  0, ;
		  0, ;
		  0, ;
		  0, ;
		  THIS.ActivityLog, ;
		  '', ;
		  '', ;
		  .F., ;
		  .F., ;
		  DATETIME(), ;
		  .F. ;
		 )
	ENDFUNC




	* Make sure the Filetypes list is delimited with spaces
	PROCEDURE FileTypes_Assign(cFileTypes)
		THIS.FileTypes = CHRTRAN(cFileTypes, ',;', '  ')
	ENDFUNC


	* ---
	* --- Definition Table methods
	* ---
	FUNCTION CreateFileTable()
		LOCAL lSuccess
		LOCAL cSafety

		m.lSuccess = .T.

		m.cSafety = SET("SAFETY")
		SET SAFETY OFF
		IF USED(JUSTSTEM(THIS.FileTable))
			USE IN (THIS.FileTable)
		ENDIF


		TRY
			CREATE TABLE (THIS.FileTable) FREE ( ;
		 	  UniqueID C(10), ;
		 	  Folder M, ;
		 	  Filename C(100), ;
		 	  FileAction C(1), ;
			  Timestamp T NULL ;
			 )
		CATCH
			m.lSuccess = .F.
			MESSAGEBOX(ERROR_CREATEFILETABLE_LOC + CHR(10) + CHR(10) + FORCEEXT(THIS.FileTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
		ENDTRY
		
		IF m.lSuccess
			* insert a record that represents open windows
			INSERT INTO (THIS.FileTable) ( ;
			  UniqueID, ;
			  Filename, ;
			  Folder, ;
			  Timestamp, ;
			  FileAction ;
			 ) VALUES ( ;
			  "WINDOW", ;
			  OPENWINDOW_LOC, ;
			  '', ;
			  DATETIME(), ;
			  FILEACTION_NODEFINITIONS ;
			 )

			INDEX ON UniqueID TAG UniqueID
			INDEX ON Filename TAG Filename
			INDEX ON FileAction TAG FileAction

			USE IN (JUSTSTEM(THIS.FileTable))
		ENDIF
		
		SET Safety &cSafety

		RETURN m.lSuccess
	ENDFUNC

	FUNCTION CreateDefTable()
		LOCAL lSuccess
		LOCAL cSafety

		m.lSuccess = .T.

		m.cSafety = SET("SAFETY")
		SET SAFETY OFF
		IF USED(JUSTSTEM(THIS.DefTable))
			USE IN (THIS.DefTable)
		ENDIF


		TRY
			CREATE TABLE (THIS.DefTable) FREE ( ;
		 	  UniqueID C(10), ;
		 	  DefType C(1), ;
		 	  FileID C(10), ;
			  Symbol M, ;
			  ClassName M, ;
			  ProcName M, ;
			  ProcLineNo I, ;
			  LineNo I, ;
			  Abstract M, ;
			  Inactive L ;
			 )
			 
		CATCH
			m.lSuccess = .F.
			MESSAGEBOX(ERROR_CREATEDEFTABLE_LOC + CHR(10) + CHR(10) + FORCEEXT(THIS.DefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
		ENDTRY

		IF m.lSuccess 
			INDEX ON UniqueID TAG UniqueID
			INDEX ON DefType TAG DefType
			INDEX ON Inactive TAG Inactive
			INDEX ON FileID TAG FileID

			USE IN (JUSTSTEM(THIS.DefTable))
		ENDIF
		
		SET Safety &cSafety

		RETURN m.lSuccess
	ENDFUNC


	* Open a File & Definition tables
	FUNCTION OpenTables(lExclusive, lQuiet)
		LOCAL lSuccess

		THIS.CloseTables()

		m.lSuccess = .T.

		IF !FILE(FORCEEXT(THIS.FileTable, "DBF"))
			m.lSuccess = THIS.CreateFileTable()
		ENDIF
		IF m.lSuccess AND !FILE(FORCEEXT(THIS.DefTable, "DBF"))
			m.lSuccess = THIS.CreateDefTable()
		ENDIF

		IF m.lSuccess
			* Open the table of files processed
			TRY
				IF m.lExclusive
					USE (THIS.FileTable) ALIAS FileCursor IN 0 EXCLUSIVE
				ELSE
					USE (THIS.FileTable) ALIAS FileCursor IN 0 SHARED AGAIN
				ENDIF
			CATCH
			ENDTRY

			* Open the table of Definitions
			TRY
				IF m.lExclusive
					USE (THIS.DefTable) ALIAS FoxDefCursor IN 0 EXCLUSIVE
				ELSE
					USE (THIS.DefTable) ALIAS FoxDefCursor IN 0 SHARED AGAIN
				ENDIF
			CATCH
			ENDTRY



			IF USED("FileCursor")
				IF TYPE("FileCursor.UniqueID") <> 'C' OR TYPE("FileCursor.Filename") <> 'C'
					m.lSuccess = .F.
					IF !m.lQuiet
						MESSAGEBOX(ERROR_BADFILETABLE_LOC + CHR(10) + CHR(10) + FORCEEXT(THIS.FileTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
					ENDIF
				ENDIF
			ELSE
				IF !m.lQuiet
					MESSAGEBOX(ERROR_OPENFILETABLE_LOC + CHR(10) + CHR(10) + FORCEEXT(THIS.FileTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
				ENDIF
				m.lSuccess = .F.
			ENDIF
			
			IF m.lSuccess
				IF USED("FoxDefCursor")
					IF TYPE("FoxDefCursor.DefType") <> 'C'
						m.lSuccess = .F.
						IF !m.lQuiet
							MESSAGEBOX(ERROR_BADDEFTABLE_LOC + CHR(10) + CHR(10) + FORCEEXT(THIS.DefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
						ENDIF
					ENDIF
				ELSE
					IF !m.lQuiet
						MESSAGEBOX(ERROR_OPENDEFTABLE_LOC + CHR(10) + CHR(10) + FORCEEXT(THIS.DefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
					ENDIF
					m.lSuccess = .F.
				ENDIF
			ENDIF

		ENDIF
		
		RETURN m.lSuccess
	ENDFUNC

	FUNCTION GetAvailableOptions()
		LOCAL oOptions
		LOCAL nSelect
		LOCAL oException
		LOCAL oRefOption
		
		nSelect = SELECT()

		oOptions = .NULL.
		oException = .NULL.
		TRY
			SELECT * FROM ;
			  FoxRefOption ;
			 ORDER BY DisplayOrd ;
			 INTO CURSOr FoxRefOptionCursor
		CATCH TO oException
		ENDTRY

		IF ISNULL(oException)
			oOptions = NEWOBJECT("Collection")
			SELECT FoxRefOptionCursor
			SCAN ALL
				oRefOption = NEWOBJECT("RefOption", "foxrefengine.prg")
				WITH oRefOption
					.OptionName   = RTRIM(FoxRefOptionCursor.OptionName)
					.Description  = FoxRefOptionCursor.Descrip
					.PropertyName = RTRIM(FoxRefOptionCursor.OptionProp)
					TRY
						.OptionValue  = EVAL("THIS." + .PropertyName)
					CATCH
					ENDTRY
				ENDWITH				
				oOptions.Add(oRefOption, oRefOption.PropertyName)
			ENDSCAN
		ELSE
			MessageBox(m.oException.Message, MB_ICONSTOP, APPNAME_LOC)
		ENDIF

		IF USED("FoxRefOption")
			USE IN FoxRefOption
		ENDIF
		IF USED("FoxRefOptionCursor")
			USE IN FoxRefOptionCursor
		ENDIF

		SELECT (nSelect)
		
		RETURN oOptions
	ENDFUNC


	FUNCTION SaveOptions(oOptionsCollection AS Collection)
		LOCAL oOptions
		LOCAL oException
		LOCAL oRefOption
		

		FOR EACH oRefOption IN oOptionsCollection
			IF !ISNULL(oRefOption.OptionValue)
				STORE (oRefOption.OptionValue) TO ("THIS." + oRefOption.PropertyName)
			ENDIF
		ENDFOR
	ENDFUNC


	FUNCTION CloseTables()
		IF USED("FoxDefCursor")
			USE IN FoxDefCursor
		ENDIF		
		IF USED("FileCursor")
			USE IN FileCursor
		ENDIF		
		IF USED("AddInCursor")
			USE IN AddInCursor
		ENDIF		
	ENDFUNC

	* Get the name of the first valid Reference table
	* We check the structure here of existing files
	* before we overwrite or open a reference table.
	* Increment a file count as necessary, so instead
	* of opening "giftrap_ref.dbf", we might end up
	* opening "giftrap_ref3.dbf" (if giftrap_ref, giftrap_ref1, and
	* giftrap_ref2 are not of the correct type)
	FUNCTION GetRefTableName(cRefTable)
		LOCAL lTableIsOkay
		LOCAL nFileCnt
		LOCAL cNewRefTable
		LOCAL oException
		LOCAL cFilename
		
		m.cFilename = CHRTRAN(JUSTFNAME(m.cRefTable), INVALID_ALIAS_CHARS, REPLICATE('_', LENC(INVALID_ALIAS_CHARS)))
		m.cRefTable = ADDBS(JUSTPATH(m.cRefTable)) + m.cFilename

		m.cNewRefTable = FORCEEXT(m.cRefTable, "DBF")
		
		m.nFileCnt = 0
		m.lTableIsOkay = .F.
		DO WHILE !lTableIsOkay AND m.nFileCnt < 100
			TRY
				IF FILE(m.cNewRefTable)
					USE (m.cNewRefTable) ALIAS CheckRefTable IN 0 SHARED AGAIN
					m.lTableIsOkay = (TYPE("CheckRefTable.RefType") == 'C')
					USE IN CheckRefTable
				ELSE
					m.lTableIsOkay = .T.
				ENDIF
			CATCH
			ENDTRY
						
			IF !m.lTableIsOkay
				m.nFileCnt = m.nFileCnt + 1
				m.cNewRefTable = ADDBS(JUSTPATH(m.cRefTable)) + JUSTSTEM(m.cRefTable) + TRANSFORM(m.nFileCnt) + ".DBF"
			ENDIF
		ENDDO
		IF m.nFileCnt > 0
			m.cRefTable = ADDBS(JUSTPATH(m.cRefTable)) + JUSTSTEM(m.cRefTable) + TRANSFORM(m.nFileCnt) + ".DBF"
		ENDIF

		RETURN m.cRefTable
	ENDFUNC

	* ---
	* --- Reference Table methods
	* ---
	
	
	FUNCTION CreateRefTable(cRefTable)
		LOCAL lSuccess
		LOCAL cSafety
		LOCAL oException

		m.lSuccess = .T.

		THIS.RefTable = ''

		m.cRefTable = THIS.GetRefTableName(m.cRefTable)
		
		m.cSafety = SET("SAFETY")
		SET SAFETY OFF

		IF USED(JUSTSTEM(m.cRefTable))
			USE IN (JUSTSTEM(m.cRefTable))
		ENDIF

		TRY
			CREATE TABLE (m.cRefTable) FREE ( ;
		 	  UniqueID C(10), ;
			  SetID C(10), ;
			  RefID C(10), ;
		 	  RefType C(1), ;
		 	  FindType C(1), ;
		 	  FileID C(10), ;
			  Symbol M, ;
			  ClassName M, ;
			  ProcName M, ;
			  ProcLineNo I, ;
			  LineNo I, ;
			  ColPos I, ;
			  MatchLen I, ;
			  Abstract M, ;
			  RecordID C(10), ;
			  UpdField C(15), ;
			  Checked L NULL, ;
			  NoReplace L, ;
			  TimeStamp T NULL, ;
			  Inactive L ;
			 )
		CATCH TO oException
			m.lSuccess = .F.
			MESSAGEBOX(ERROR_CREATEREFTABLE_LOC + " (" + m.oException.Message + "):" + CHR(10) + CHR(10) + FORCEEXT(m.cRefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
		ENDTRY
		
		IF m.lSuccess
			INDEX ON RefType TAG RefType
			INDEX ON SetID TAG SetID
			INDEX ON RefID TAG RefID
			INDEX ON UniqueID TAG UniqueID
			INDEX ON FileID TAG FileID
			INDEX ON Checked TAG Checked
			INDEX ON Inactive TAG Inactive


			* add the record that holds our results window search position & other options
			INSERT INTO (m.cRefTable) ( ;
			  UniqueID, ;
			  SetID, ;
			  RefType, ;
			  FindType, ;
			  FileID, ;
			  Symbol, ;
			  ClassName, ;
			  ProcName, ;
			  ProcLineNo, ;
			  LineNo, ;
			  ColPos, ;
			  MatchLen, ;
			  Abstract, ;
			  RecordID, ;
			  UpdField, ;
			  Timestamp, ;
			  Checked, ;
			  NoReplace, ;
			  Inactive ;
			 ) VALUES ( ;
			  SYS(2015), ;
			  '', ;
			  REFTYPE_INIT, ;
			  '', ;
			  '', ;
			  THIS.ProjectFile, ;
			  '', ;
			  '', ;
			  0, ;
			  0, ;
			  0, ;
			  0, ;
			  '', ;
			  '', ;
			  '', ;
			  DATETIME(), ;
			  .F., ;
			  .F., ;
			  .F. ;
			 )

			THIS.RefTable = m.cRefTable

			USE
		ENDIF
		
		SET SAFETY &cSafety

		RETURN m.lSuccess
	ENDFUNC
	
	* Open a FoxRef table 
	* Return TRUE if table exists and it's in the correct format
	* [lCreate]    = True to create table if it doesn't exist
	* [lExclusive] = True to open for exclusive use
	FUNCTION OpenRefTable(cRefTable, lExclusive)
		LOCAL lSuccess
		LOCAL oException

		IF USED("FoxRefCursor")
			USE IN FoxRefCursor
		ENDIF		
		THIS.RefTable = ''

		m.lSuccess = .T.

		m.cRefTable = THIS.GetRefTableName(m.cRefTable)

		IF !FILE(FORCEEXT(m.cRefTable, "DBF"))
			m.lSuccess = THIS.CreateRefTable(m.cRefTable)
		ENDIF

		IF m.lSuccess
			TRY
				IF m.lExclusive
					USE (m.cRefTable) ALIAS FoxRefCursor IN 0 EXCLUSIVE
				ELSE
					USE (m.cRefTable) ALIAS FoxRefCursor IN 0 SHARED AGAIN
				ENDIF
			CATCH TO oException
				MESSAGEBOX(ERROR_OPENREFTABLE_LOC + " (" + m.oException.Message + "):" + CHR(10) + CHR(10) + FORCEEXT(m.cRefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
				m.lSuccess = .F.
			ENDTRY

			IF m.lSuccess
				IF TYPE("FoxRefCursor.RefType") == 'C'
					THIS.RefTable = m.cRefTable
				ELSE
					m.lSuccess = .F.
					MESSAGEBOX(ERROR_BADREFTABLE_LOC + CHR(10) + CHR(10) + FORCEEXT(m.cRefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
				ENDIF
			ENDIF
		ENDIF
		
		RETURN m.lSuccess
	ENDFUNC

	* return number of search sets in the current FoxRef table
	FUNCTION SearchCount()
		LOCAL nSelect
		LOCAL nSearchCnt
		LOCAL ARRAY aSearchCnt[1]
		
		m.nSelect = SELECT()
		SELECT CNT(*) ;
		 FROM (THIS.RefTable) ;
		 WHERE ;
		  RefType == REFTYPE_SEARCH AND !Inactive ;
		 INTO ARRAY aSearchCnt
		IF _TALLY > 0
			m.nSearchCnt = aSearchCnt[1]
		ELSE
			m.nSearchCnt = 0
		ENDIF
		
		SELECT (m.nSelect)
		
		RETURN m.nSearchCnt
	ENDFUNC
	
	* Returns TRUE if there is only 1 record in the FoxRef table, 
	* indicating that no searches have been done yet
	* (the first record is initialization information)
	FUNCTION FirstSearch()
		RETURN USED("FoxRefCursor") AND RECCOUNT("FoxRefCursor") <= 1
	ENDFUNC

	* Abstract:
	*   Set a specific project to display result sets for.
	*	Pass an empty string or "global" to display result sets
	*	that are not associated with a project.
	*
	* Parameters:
	*   [cProject]
	FUNCTION SetProject(cProjectFile, lOverwrite)
		LOCAL lSuccess
		LOCAL cRefTable
		LOCAL i
		LOCAL lFoundProject
		LOCAL lOpened
		LOCAL oFileRef
		LOCAL oErr

		lSuccess = .F.

		IF VARTYPE(cProjectFile) <> 'C' 
			cProjectfile = THIS.ProjectFile
		ENDIF

		THIS.oProjectFileRef = .NULL.
		IF EMPTY(cProjectFile)
			* use the active project if a project name is not passsed
			IF Application.Projects.Count > 0
				cProjectFile = Application.ActiveProject.Name
			ELSE
				cProjectFile = PROJECT_GLOBAL
			ENDIF
			
			lSuccess = .T.
		ELSE
			* make sure Project specified is open
			IF cProjectFile == PROJECT_GLOBAL
				lSuccess = .T.
			ELSE
				cProjectFile = UPPER(FORCEEXT(FULLPATH(cProjectFile), "PJX"))

				FOR i = 1 TO Application.Projects.Count
					IF UPPER(Application.Projects(i).Name) == cProjectFile
						cProjectFile = Application.Projects(i).Name
						lSuccess = .T.
						EXIT
					ENDIF
				ENDFOR
				IF !lSuccess
					IF FILE(cProjectFile)
						lOpened = .T.
						* open the project
						TRY
							MODIFY PROJECT (cProjectFile) NOWAIT

						CATCH TO oErr
							MESSAGEBOX(oErr.Message, MB_ICONEXCLAMATION, APPNAME_LOC)
							lOpened = .F.
						ENDTRY

						IF lOpened
							* search again to find where in the Projects collection it is
							FOR i = 1 TO Application.Projects.Count
								IF UPPER(Application.Projects(i).Name) == cProjectFile
									cProjectFile = Application.Projects(i).Name
									lSuccess = .T.
									EXIT
								ENDIF
							ENDFOR
						ENDIF
					ENDIF
				ENDIF
				
			ENDIF
		ENDIF

		IF lSuccess
			IF EMPTY(cProjectFile) OR cProjectFile == PROJECT_GLOBAL
				THIS.ProjectFile = PROJECT_GLOBAL
				cRefTable        = THIS.FoxRefDirectory + GLOBAL_TABLE + RESULT_EXT
			ELSE
				THIS.ProjectFile = UPPER(cProjectFile)
				cRefTable        = THIS.GetRefTableName(ADDBS(JUSTPATH(cProjectFile)) + JUSTSTEM(cProjectFile) + RESULT_EXT)

				THIS.oProjectFileRef = Application.ActiveProject
			ENDIF

			IF lOverwrite
				lSuccess = THIS.CreateRefTable(cRefTable)
			ELSE
				lSuccess = THIS.OpenRefTable(cRefTable)
			ENDIF
		ENDIF
		RETURN lSuccess
	ENDFUNC

	* Add to cursor of available project files.
	* This cursor is all files in current project,
	* plus any files we encounter along the way
	* that are #include or SET PROCEDURE TO, SET CLASSLIB TO
	FUNCTION AddFileToProjectCursor(cFileName)
		LOCAL nSelect
		LOCAL i
		LOCAL nCnt
		LOCAL ARRAY aFileList[1]
		
		m.nSelect = SELECT()
		
		m.cFilename = LOWER(FULLPATH(m.cFilename))
		m.cFolder   = JUSTPATH(m.cFilename)

		m.cFilename = PADR(JUSTFNAME(m.cFilename), 100)

		SELECT ProjectFilesCursor
		LOCATE FOR Filename == m.cFilename AND Folder == m.cFolder
		IF !FOUND()
			INSERT INTO ProjectFilesCursor ( ;
			  Filename, ;
			  Folder ;
			 ) VALUES ( ;
			  m.cFilename, ;
			  m.cFolder ;
			 )
		ENDIF

		SELECT (m.nSelect)
	ENDFUNC
	
	* Grabs all Include files for this project
	* and adds them to our project list cursor
	FUNCTION UpdateProjectFiles()
		LOCAL nSelect
		LOCAL i
		LOCAL nCnt
		LOCAL oFileRef
		LOCAL ARRAY aFileList[1]
		
		m.nSelect = SELECT()
		
		IF USED("ProjectFilesCursor")
			USE IN ProjectFilesCursor
		ENDIF
		CREATE CURSOR ProjectFilesCursor ( ;
		  Folder M, ;
		  Filename C(100) ;
		 )
		INDEX ON Filename TAG Filename

		* Add in all files that are in currently in the project
		IF VARTYPE(THIS.oProjectFileRef) == 'O'
			FOR EACH m.oFileRef IN THIS.oProjectFileRef.Files
				THIS.AddFileToProjectCursor(oFileRef.Name)
			ENDFOR
		ENDIF

		* add in all dependencies
		SELECT ;
		  DISTINCT PADR(LEFTC(Symbol, 254), 254) AS IncludeFile ;
		 FROM (THIS.DefTable) DefTable ;
		 WHERE ;
		  DefTable.DefType == DEFTYPE_INCLUDEFILE AND ;
		  !DefTable.Inactive ;
		 INTO ARRAY aFileList
		 m.nCnt = _TALLY
		 FOR m.i = 1 TO m.nCnt
		 	TRY
			 	IF FILE(RTRIM(aFileList[m.i, 1]))  && make sure we can find the file along our path somewhere
				 	THIS.AddFileToProjectCursor(aFileList[m.i, 1])
				ENDIF
			CATCH
			ENDTRY
		 ENDFOR

		SELECT (m.nSelect)
	ENDFUNC

	* return project files as a collection
	FUNCTION GetProjectFiles()
		LOCAL nSelect
		LOCAL oProjectFiles
		
		nSelect = SELECT()
		
		oProjectFiles = CREATEOBJECT("Collection")
		
		IF USED("ProjectFilesCursor")
			SELECT ProjectFilesCursor
			SCAN ALL
				oProjectFiles.Add(ADDBS(RTRIM(ProjectFilesCursor.Folder)) + RTRIM(ProjectFilesCursor.FileName))
			ENDSCAN
		ENDIF		
		SELECT (nSelect)
		
		RETURN oProjectFiles
	ENDFUNC
	
	* collect all definitions for a project/folder and currently open window
	* without doing an actual search
	* [lLocalOnly] = search open window for LOCALS & PARAMETERS
	FUNCTION CollectDefinitions(lLocalOnly)
		LOCAL nSelect
		LOCAL cOpenFile
		LOCAL nCnt
		LOCAL i
		LOCAL lSuccess
		LOCAL lDefinitionsOnly
		LOCAL lOverwritePrior
		LOCAL lCodeOnly
		LOCAL cFileTypes
		LOCAL lProjectHomeDir
		LOCAL oException
		

		IF !THIS.SearchInit()
			RETURN .F.
		ENDIF

		lDefinitionsOnly  = THIS.lDefinitionsOnly
		lOverwritePrior   = THIS.OverwritePrior
		lCodeOnly         = THIS.CodeOnly
		cFileTypes        = THIS.FileTypes
		lProjectHomeDir   = THIS.ProjectHomeDir
		
		THIS.lDefinitionsOnly = .T.
		THIS.OverwritePrior   = .F.
		THIS.CodeOnly         = .F.
		THIS.FileTypes        = FILETYPES_DEFINITIONS
		THIS.ProjectHomeDir   = THIS.AutoProjectHomeDir

		m.nSelect = SELECT()

		* make sure we have the latest definitions for the open window
		TRY
			IF m.lLocalOnly
				UPDATE FoxDefCursor ;
				 SET Inactive = .T. ;
				 WHERE FileID = "WINDOW"
				IF THIS.WindowHandle >= 0 AND VARTYPE(THIS.oWindowEngine) == 'O'
					THIS.oWindowEngine.WindowHandle = THIS.WindowHandle
					THIS.oWindowEngine.ProcessDefinitions(THIS)

					UPDATE FileCursor SET ;
					  Filename = JUSTFNAME(THIS.WindowFilename), ;
					  Folder = JUSTPATH(THIS.WindowFilename), ;
					  FileAction = FILEACTION_DEFINITIONS ;
					 WHERE UniqueID = "WINDOW"
				ENDIF
			ELSE
				IF THIS.SetProject(.NULL., .F.)
					IF THIS.ProjectFile == PROJECT_GLOBAL OR EMPTY(THIS.ProjectFile)
						m.lSuccess = THIS.FolderSearch('')
					ELSE
						m.lSuccess = THIS.ProjectSearch('')
					ENDIF
				ENDIF

				* Process definitions for files that
				* were #included
				m.nCnt = THIS.oFileCollection.Count
				FOR m.i = 1 TO m.nCnt
					m.lSuccess = THIS.FileSearch(THIS.oFileCollection.Item(m.i), '')
					IF !m.lSuccess
						EXIT
					ENDIF
				ENDFOR
				THIS.oFileCollection.Remove(-1)
				THIS.oProcessedCollection.Remove(-1)

				THIS.UpdateProjectFiles()
			ENDIF
		CATCH TO oException
			MESSAGEBOX(oException.Message)
		ENDTRY
		
		THIS.lDefinitionsOnly   = lDefinitionsOnly
		THIS.OverwritePrior     = lOverwritePrior
		THIS.CodeOnly           = lCodeOnly
		THIS.FileTypes          = cFileTypes
		THIS.ProjectHomeDir     = lProjectHomeDir

		SELECT (m.nSelect)
	ENDFUNC


	FUNCTION Search(cPattern, lShowDialog)
		** No search
		
		RETURN .F.
	ENDFUNC


	* Determine if the Reference table we want to open is 
	* actually one of ours.  If we're overwriting or a reference
	* table doesn't exist for this project, then create a new 
	* Reference Table.
	*
	* Once we have a reference table, then we add a new record
	* that represents the search criteria for this particular
	* search.
	FUNCTION UpdateRefTable(cScope, cPattern, cProjectOrDir)
		LOCAL nSelect
		LOCAL cSafety
		LOCAL cSearchOptions
		LOCAL cRefTable
		LOCAL i
		
		m.nSelect = SELECT()
		
		IF VARTYPE(cRefTable) <> 'C' OR EMPTY(cRefTable)
			cRefTable = THIS.RefTable
		ENDIF

		IF EMPTY(cRefTable)
			RETURN .F.
		ENDIF
		
		IF USED("FoxRefCursor")
			USE IN FoxRefCursor
		ENDIF

		IF !THIS.OpenRefTable(cRefTable)
			RETURN .F.
		ENDIF


		THIS.tTimeStamp = DATETIME()


		* Since we're only doing definitions and there
		* is no search in progress, then don't create
		* a Search Set record in the FoxRef table
		IF THIS.lDefinitionsOnly
			RETURN .T.
		ENDIF

		* build a string representing the search options that
		* we can store to the FoxRef cursor
		cSearchOptions = IIF(THIS.Comments == COMMENTS_EXCLUDE, 'X', '') + ;
		                 IIF(THIS.Comments == COMMENTS_ONLY, 'C', '') + ;
		                 IIF(THIS.MatchCase, 'M', '') + ;
		                 IIF(THIS.WholeWordsOnly, 'W', '') + ;
		                 IIF(THIS.ProjectHomeDir, 'H', '') + ;
		                 IIF(THIS.FormProperties, 'P', '') + ;
		                 IIF(THIS.SubFolders, 'S', '') + ;
		                 IIF(THIS.Wildcards, 'Z', '') + ;
		                 ';' + ALLTRIM(THIS.FileTypes)


		* if we've already searched for this same exact symbol
		* with the same exact criteria in the same exact project/folder,
		* then simply update what we have
		IF THIS.lRefreshMode
			SELECT FoxRefCursor
			LOCATE FOR SetID == THIS.cSetID AND RefType == REFTYPE_SEARCH AND !Inactive
			THIS.lRefreshMode = FOUND()
		ENDIF
		
		IF !THIS.lRefreshMode
			SELECT FoxRefCursor
	 		LOCATE FOR RefType == REFTYPE_SEARCH AND ClassName == cProjectOrDir AND Symbol == cPattern AND Abstract == cSearchOptions AND !Inactive
			THIS.lRefreshMode = FOUND()
		ENDIF

		IF THIS.lRefreshMode
			THIS.tTimeStamp = FoxRefCursor.TimeStamp

			THIS.cSetID = FoxRefCursor.SetID
			UPDATE FoxRefCursor ;
			 SET Inactive = .T. ;
			 WHERE ;
			  SetID == THIS.cSetID AND ;
			  (RefType == REFTYPE_RESULT OR RefType == REFTYPE_ERROR OR RefType == REFTYPE_NOMATCH)
		ELSE
			THIS.cSetID = SYS(2015)
	
			* add the record that specifies the search criteria, etc
			INSERT INTO FoxRefCursor ( ;
			  UniqueID, ;
			  SetID, ;
			  RefType, ;
			  FindType, ;
			  Symbol, ;
			  ClassName, ;
			  ProcName, ;
			  ProcLineNo, ;
			  LineNo, ;
			  ColPos, ;
			  MatchLen, ;
			  Abstract, ;
			  RecordID, ;
			  UpdField, ;
			  Checked, ;
			  NoReplace, ;
			  Timestamp, ;
			  Inactive ;
			 ) VALUES ( ;
			  SYS(2015), ;
			  THIS.cSetID, ;
			  REFTYPE_SEARCH, ;
			  '', ;
			  cPattern, ;
  			  cProjectOrDir, ;
			  '', ;
			  0, ;
			  0, ;
			  0, ;
			  0, ;
			  cSearchOptions, ;
			  '', ;
			  '', ;
			  .F., ;
			  .F., ;
			  DATETIME(), ;
			  .F. ;
			 )
		ENDIF

		* update each of the search engines with the new SetID
		IF VARTYPE(THIS.oWindowEngine) == 'O'
			THIS.oWindowEngine.SetID = THIS.cSetID
		ENDIF
		FOR m.i = 1 TO THIS.oEngineCollection.Count
			IF VARTYPE(THIS.oEngineCollection.Item(m.i)) == 'O'
				WITH THIS.oEngineCollection.Item(m.i)
					.SetID = THIS.cSetID
				ENDWITH
			ENDIF
		ENDFOR

		SELECT DISTINCT SetID, FileID, TimeStamp ;
		 FROM FoxRefCursor ;
		 WHERE (RefType == REFTYPE_RESULT OR RefType == REFTYPE_NOMATCH) AND !Inactive ;
		 INTO CURSOR FoxRefSearchedCursor

		
		SELECT (m.nSelect)
		
		RETURN .T.
	ENDFUNC


	* -- Search a Folder
	FUNCTION FolderSearch(cPattern, cFileDir)
		LOCAL i, j
		LOCAL cFileDir
		LOCAL nFileTypesCnt
		LOCAL cFileTypes
		LOCAL lAutoYield
		LOCAL lSuccess
		LOCAL ARRAY aFileList[1]
		LOCAL ARRAY aFileTypes[1]

		IF VARTYPE(cPattern) <> 'C'
			cPattern = THIS.Pattern
		ENDIF

		IF VARTYPE(cFileDir) <> 'C' OR EMPTY(cFileDir)
			cFileDir = ADDBS(THIS.FileDirectory)
		ELSE
			cFileDir = ADDBS(cFileDir)
		ENDIF

		IF EMPTY(cFileDir) OR !DIRECTORY(cFileDir)
			RETURN .F.
		ENDIF

		IF !THIS.SearchInit()
			RETURN .F.
		ENDIF

		IF !THIS.UpdateRefTable(SCOPE_FOLDER, cPattern, cFileDir)
			RETURN .F.
		ENDIF
		
		cFileTypes = CHRTRAN(THIS.FileTypes, ',;', '  ')
		nFileTypesCnt = ALINES(aFileTypes, ALLTRIM(cFileTypes), .T., ' ')

		lAutoYield = _VFP.AutoYield
		_VFP.AutoYield = .T.


		lSuccess = THIS.ProcessFolder(cFileDir, cPattern, @aFileTypes, nFileTypesCnt)

		THIS.CloseProgress()
		
		_VFP.AutoYield = lAutoYield

		THIS.UpdateLookForMRU(cPattern)
		THIS.UpdateFolderMRU(cFileDir)
		THIS.UpdateFileTypesMRU(cFileTypes)
		
		RETURN lSuccess
	ENDFUNC

	* used in conjuction with FolderSearch() for
	* when we're searching subfolders
	FUNCTION ProcessFolder(cFileDir, cPattern, aFileTypes, nFileTypesCnt)
		LOCAL nFolderCnt
		LOCAL cFilename
		LOCAL i, j
		LOCAL nFileCnt
		LOCAL nProgress
		LOCAL lSuccess
		LOCAL ARRAY aFileList[1]
		LOCAL ARRAY aFolderList[1]

		cFileDir = ADDBS(cFileDir)

		lSuccess = .T.

		IF THIS.ShowProgress
			* determine how many files there are to process
			nFileCnt = 0
			FOR i = 1 TO nFileTypesCnt
				TRY
					nFileCnt = nFileCnt + ADIR(aFileList, cFileDir + aFileTypes[i], '', 1)
				CATCH
					* ADIR failed, so just ignore
				ENDTRY
			ENDFOR

			* Set the description on the Progress form
			THIS.ProgressInit(IIF(THIS.lRefreshMode, PROGRESS_REFRESHING_LOC, IIF(THIS.lDefinitionsOnly, PROGRESS_DEFINITIONS_LOC, PROGRESS_SEARCHING_LOC)) + ' ' + DISPLAYPATH(cFileDir, 40) + IIF(!THIS.lRefreshMode OR THIS.lDefinitionsOnly, '',  ' [' + cPattern + ']'), nFileCnt)
		ENDIF

		nProgress = 0
		FOR i = 1 TO nFileTypesCnt
			IF THIS.lCancel
				EXIT
			ENDIF

			TRY
				nFileCnt = ADIR(aFileList, cFileDir + aFileTypes[i], '', 1)
			CATCH
				nFileCnt = 0
			ENDTRY
			FOR j = 1 TO nFileCnt
				IF THIS.lCancel
					EXIT
				ENDIF

				cFilename = aFileList[j, 1]
				nProgress = nProgress + 1

				THIS.UpdateProgress(nProgress, cFilename, .T.)
				IF !THIS.FileSearch(cFileDir + cFilename, cPattern)
					EXIT
				ENDIF
			ENDFOR
		ENDFOR
		
		* Process any sub-directories
		IF !THIS.lCancel
			IF THIS.SubFolders
				TRY
					nFolderCnt = ADIR(aFolderList, cFileDir + "*.*", 'D', 1)
				CATCH
					nFolderCnt = 0
				ENDTRY
				FOR i = 1 TO nFolderCnt
					IF !aFolderList[i, 1] == '.' AND !aFolderList[i, 1] == '..' AND 'D'$aFolderList[i, 5] AND DIRECTORY(cFileDir + aFolderList[i, 1])
						THIS.ProcessFolder(cFileDir + aFolderList[i, 1], cPattern, @aFileTypes, nFileTypesCnt)
					ENDIF
					IF THIS.lCancel
						EXIT
					ENDIF
				ENDFOR
			ENDIF
		ENDIF
		
		RETURN lSuccess
	ENDFUNC

	
	* -- Search files in a Project
	* -- Pass an empty cPattern to only collect definitions
	FUNCTION ProjectSearch(cPattern, cProjectFile)
		LOCAL nFileIndex
		LOCAL nProjectIndex
		LOCAL oProjectRef
		LOCAL oFileRef
		LOCAL cFileTypes
		LOCAL nFileTypesCnt
		LOCAL lAutoYield
		LOCAL lSuccess
		LOCAL nFileCnt
		LOCAL lSuccess
		LOCAL i
		LOCAL cHomeDir
		LOCAL oDefFileTypes
		LOCAL cFilePath
		LOCAL oMatchFileCollection
		LOCAL ARRAY aFileTypes[1]
		LOCAL ARRAY aFileList[1]

		IF VARTYPE(cPattern) <> 'C'
			cPattern = THIS.Pattern
		ENDIF
		
		IF VARTYPE(cProjectFile) <> 'C' OR EMPTY(cProjectFile)
			cProjectFile = THIS.ProjectFile
		ENDIF

		IF !THIS.SearchInit()
			RETURN .F.
		ENDIF
		
		IF !THIS.UpdateRefTable(SCOPE_PROJECT, cPattern, THIS.ProjectFile)
			RETURN .F.
		ENDIF


		lSuccess = .T.

		cFileTypes = THIS.FileTypes
		
		
		oMatchFileCollection = NEWOBJECT("CFoxRefCollection", "FoxRefCollection.prg")
		oMatchFileCollection.AddList(cFileTypes, ' ')
		
		
		
		oDefFileTypes = NEWOBJECT("CFoxRefCollection", "FoxRefCollection.prg")
		oDefFileTypes.AddList(UPPER(FILETYPES_DEFINITIONS), ' ')

		lAutoYield = _VFP.AutoYield
		_VFP.AutoYield = .T.

		FOR EACH oProjectRef IN Application.Projects
			IF UPPER(oProjectRef.Name) == THIS.ProjectFile
				cHomeDir = ADDBS(UPPER(oProjectRef.HomeDir))
			
				* determine how many files to process so we can update the Progress form
				nFileCnt = 0
				IF THIS.ShowProgress
					FOR EACH oFileRef IN oProjectRef.Files
						IF (!THIS.ProjectHomeDir OR ADDBS(UPPER(JUSTPATH(oFileRef.Name))) = cHomeDir) AND ;
						     ((THIS.IncludeDefTable AND oDefFileTypes.GetIndex("*." + UPPER(JUSTEXT(oFileRef.Name))) > 0) OR THIS.ProjectMatch(cFileTypes, oMatchFileCollection, oFileRef.Name))
							nFileCnt = nFileCnt + 1
						ENDIF
					ENDFOR
					THIS.ProgressInit(IIF(THIS.lRefreshMode, PROGRESS_REFRESHING_LOC, IIF(THIS.lDefinitionsOnly, PROGRESS_DEFINITIONS_LOC, PROGRESS_SEARCHING_LOC)) + ' ' + JUSTFNAME(THIS.ProjectFile) + IIF(!THIS.lRefreshMode OR THIS.lDefinitionsOnly, '',  ' [' + cPattern + ']'), nFileCnt)
				ENDIF


				* now process each file in the project that matches our filetypes
				i = 0
				FOR EACH oFileRef IN oProjectRef.Files
					IF (!THIS.ProjectHomeDir OR ADDBS(UPPER(JUSTPATH(oFileRef.Name))) = cHomeDir)
						IF THIS.ProjectMatch(cFileTypes, oMatchFileCollection, oFileRef.Name)
							i = i + 1
							THIS.UpdateProgress(i, oFileRef.Name, .T.)
							IF !THIS.FileSearch(oFileRef.Name, cPattern)
								EXIT
							ENDIF
						ELSE
							IF THIS.IncludeDefTable AND oDefFileTypes.GetIndex("*." + UPPER(JUSTEXT(oFileRef.Name))) > 0
								* don't search the file because it doesn't match the filetypes
								* specified -- however, we still want to collection definitions
								* on this filetype
								i = i + 1
								THIS.UpdateProgress(i, oFileRef.Name, .T.)
								IF !THIS.FileSearch(oFileRef.Name, '')
									EXIT
								ENDIF
							ENDIF
						ENDIF
					ENDIF
					IF THIS.lCancel
						lSuccess = .F.
						EXIT
					ENDIF
				ENDFOR
				
				EXIT
			ENDIF
		ENDFOR

		oFileTypesCollection = .NULL.

		THIS.CloseProgress()

		_VFP.AutoYield = lAutoYield

		THIS.UpdateLookForMRU(cPattern)
		THIS.UpdateFileTypesMRU(cFileTypes)
				
		RETURN lSuccess
	ENDFUNC


	* Return Search Engine to use based upon filetype
	FUNCTION GetEngine(m.cFilename)
		LOCAL nEngineIndex
		LOCAL i
		
		IF THIS.oEngineCollection.Count == 0
			RETURN .NULL.
		ENDIF
		
		m.cFilename = UPPER(m.cFilename)
		
		* determine which search engine to use based upon the filetype
		m.nEngineIndex = 1
		FOR m.i = 2 TO THIS.oEngineCollection.Count
			IF VARTYPE(THIS.oEngineCollection.Item(m.i)) == 'O'
				IF THIS.WildcardMatch(THIS.oEngineCollection.GetKey(m.i), m.cFilename)
					m.nEngineIndex = m.i
					EXIT
				ENDIF
			ENDIF
		ENDFOR
		
		* if we're still using the default search engine,
		* then make sure this file type isn't set to 
		* be excluded
		IF m.nEngineIndex == 1
			FOR m.i = 2 TO THIS.oEngineCollection.Count
				IF VARTYPE(THIS.oEngineCollection.Item(m.i)) <> 'O'
					IF THIS.WildcardMatch(THIS.oEngineCollection.GetKey(m.i), m.cFilename)
						m.nEngineIndex = m.i
						EXIT
					ENDIF
				ENDIF
			ENDFOR
		ENDIF
		
		RETURN THIS.oEngineCollection.Item(m.nEngineIndex)
	ENDFUNC


	* Search a file
	* -- Pass an empty cPattern to only collect definitions
	FUNCTION FileSearch(cFilename, cPattern)
		LOCAL nSelect
		LOCAL cFileFind
		LOCAL cFolderFind
		LOCAL nSelect
		LOCAL lDefinitions
		LOCAL lSearch
		LOCAL cFileID
		LOCAL cFileAction
		LOCAL oEngine
		LOCAL lSuccess
		LOCAL nIndex
		LOCAL cLowerFilename
		LOCAL ARRAY aFileList[1]

		IF THIS.lCancel
			RETURN .F.
		ENDIF
		
		* make sure we haven't searched for this file already
		m.cLowerFilename = LOWER(m.cFilename)
		FOR m.nIndex = 1 TO THIS.oProcessedCollection.Count
			IF THIS.oProcessedCollection.Item(m.nIndex) == m.cLowerFilename
				RETURN .T.
			ENDIF
		ENDFOR
		
		THIS.oProcessedCollection.Add(LOWER(m.cFilename))

		IF VARTYPE(m.cPattern) <> 'C'
			m.cPattern = THIS.Pattern
		ENDIF

		m.nSelect = SELECT()

		m.lSearch      = !EMPTY(m.cPattern)
		m.lDefinitions = !m.lSearch OR THIS.IncludeDefTable && false to search only

		m.oEngine = THIS.GetEngine(JUSTFNAME(m.cFilename))

		m.lSuccess = .T.

		* if we don't have an engine object defined, then assume
		* we're supposed to ignore this filetype
		IF VARTYPE(m.oEngine) == 'O'
			WITH m.oEngine
				.Filename   = m.cFilename
				.SetTimeStamp()

				m.cFileFind   = PADR(LOWER(JUSTFNAME(m.cFilename)), 100)
				* m.cFolderFind = PADR(LOWER(JUSTPATH(m.cFilename)), 240)
				m.cFolderFind = LOWER(JUSTPATH(m.cFilename))

				SELECT FileCursor
				LOCATE FOR FileName == m.cFileFind AND Folder == m.cFolderFind
				IF FOUND()
					* This file has been previously processed
					* However, if the TimeStamp is different or we've
					* never processed definitions, then we need to reprocess
					m.cFileID  = FileCursor.UniqueID
					IF m.lDefinitions
						m.lDefinitions = (FileCursor.TimeStamp <> .FileTimeStamp) OR FileCursor.FileAction <> FILEACTION_DEFINITIONS OR m.oEngine.NoRefresh
					ENDIF

					REPLACE ;
					  TimeStamp WITH .FileTimeStamp ;
					 IN FileCursor

					IF m.lDefinitions
						UPDATE FoxDefCursor SET Inactive = .T. WHERE FileID == m.cFileID
					ENDIF
				ELSE
					* file has never been processed, so add a file record to
					* the RefFile table
					m.cFileID  = SYS(2015)
	
					INSERT INTO FileCursor ( ;
					  UniqueID, ;
					  Filename, ;
					  Folder, ;
					  Timestamp, ;
					  FileAction ;
					 ) VALUES ( ;
					  m.cFileID, ;
					  m.cFileFind, ;
					  m.cFolderFind, ;
					  .FileTimeStamp, ;
					  FILEACTION_NODEFINITIONS ;
					 )
				ENDIF

				* determine whether or not we need to process
				* this file at all.  Two conditions that might 
				* force us to process:
				*  1) We need to collect definitions
				*  2) We need to search file for references (only if cPattern is not empty)
				IF m.lSearch
					SELECT FoxRefSearchedCursor
					LOCATE FOR SetID == THIS.cSetID AND FileID == m.cFileID
					IF FOUND()
						m.lSearch = FoxRefSearchedCursor.TimeStamp <> FileCursor.TimeStamp
						IF !m.lSearch
							UPDATE FoxRefCursor ;
							 SET Inactive = .F. ;
							 WHERE ;
							  SetID == THIS.cSetID AND ;
							  FileID == m.cFileID AND ;
							  (RefType == REFTYPE_RESULT OR RefType == REFTYPE_NOMATCH) AND ;
							  Inactive
						ENDIF
					ELSE
						m.lSearch = .T.
					ENDIF
				ENDIF

				IF m.lSearch OR m.lDefinitions
					.FileID = m.cFileID

					m.cFileAction = .SearchFor(m.cPattern, m.lSearch, m.lDefinitions, THIS)
					
					m.lSuccess = !(m.cFileAction == FILEACTION_STOP)

					IF m.lSuccess
						IF !EMPTY(m.cFileAction)
							REPLACE FileAction WITH m.cFileAction IN FileCursor
						ENDIF
					ELSE
						MESSAGEBOX(ERROR_SEARCHENGINE_LOC, MB_ICONEXCLAMATION, APPNAME_LOC)
					ENDIF
				ENDIF
			ENDWITH

			IF m.lSearch
			 	UPDATE FoxRefCursor ;
			 	  SET RefType = REFTYPE_INACTIVE ;
				  WHERE ;
				    SetID == THIS.cSetID AND ;
				    FileID == m.cFileID AND ;
				    Inactive
			ENDIF
		ENDIF

		SELECT (m.nSelect)
		
		RETURN m.lSuccess
	ENDFUNC
	

	
	* refresh results for all Sets in the Ref table or a single set
	FUNCTION RefreshResults(cSetID)
		LOCAL nSelect
		LOCAL i
		LOCAL nCnt
		LOCAL lSuccess
		LOCAL ARRAY aRefList[1]

		m.nSelect = SELECT()

		m.lSuccess = .T.

		IF VARTYPE(m.cSetID) == 'C' AND !EMPTY(m.cSetID)
			THIS.RefreshResultSet(m.cSetID)
		ELSE
			IF FILE(FORCEEXT(THIS.RefTable, "dbf"))
				m.nSelect = SELECT()

				m.lSuccess = THIS.OpenRefTable(THIS.RefTable)
				IF m.lSuccess
					SELECT SetID ;
					 FROM FoxRefCursor ;
					 WHERE RefType == REFTYPE_SEARCH AND !Inactive ;
					 INTO ARRAY aRefList
					m.nCnt = _TALLY

					FOR m.i = 1 TO m.nCnt
						THIS.RefreshResultSet(aRefList[i])
						IF THIS.lCancel
							EXIT
						ENDIF
					ENDFOR
				ENDIF


				SELECT (m.nSelect)
			ENDIF
			THIS.cSetID = ''
		ENDIF
		
		RETURN m.lSuccess
	ENDFUNC

	* refresh an existing search set
	FUNCTION RefreshResultSet(cSetID)
		LOCAL nSelect
		LOCAL lSuccess
		LOCAL cScope
		LOCAL cFolder
		LOCAL cProject
		LOCAL cPattern
		LOCAL cSearchOptions

		m.lSuccess = .F.

		IF FILE(FORCEEXT(THIS.RefTable, "dbf"))
			m.nSelect = SELECT()

			IF !THIS.OpenRefTable(THIS.RefTable)
				RETURN .F.
			ENDIF
			
			THIS.cSetID = m.cSetID
		
			SELECT FoxRefCursor
			LOCATE FOR RefType == REFTYPE_SEARCH AND SetID == m.cSetID AND !Inactive
			lSuccess = FOUND()
			IF lSuccess
				cSearchOptions = LEFTC(FoxRefCursor.Abstract, AT_C(';', FoxRefCursor.Abstract) - 1)
				IF 'X'$cSearchOptions
					THIS.Comments = COMMENTS_EXCLUDE
				ENDIF
				IF 'C'$cSearchOptions
					THIS.Comments = COMMENTS_ONLY
				ENDIF
				THIS.MatchCase      = 'M' $ cSearchOptions
				THIS.WholeWordsOnly = 'W' $ cSearchOptions
				THIS.FormProperties = 'P' $ cSearchOptions
				THIS.ProjectHomeDir = 'H' $ cSearchOptions
				THIS.SubFolders     = 'S' $ cSearchOptions
				THIS.Wildcards      = 'Z' $ cSearchOptions

				THIS.OverwritePrior = .F.

				THIS.FileTypes = ALLTRIM(SUBSTRC(FoxRefCursor.Abstract, AT_C(';', FoxRefCursor.Abstract) + 1))

				cFolder  = RTRIM(FoxRefCursor.ClassName)
				cProject = ''

				IF UPPER(JUSTEXT(cFolder)) == "PJX"
					cScope = SCOPE_PROJECT
					cProject = cFolder
				ELSE
					cScope = SCOPE_FOLDER
				ENDIF
				
				cPattern = FoxRefCursor.Symbol
			ENDIF


			IF lSuccess
				DO CASE
				CASE cScope == SCOPE_FOLDER
					lSuccess = THIS.FolderSearch(cPattern, cFolder)
				CASE cScope == SCOPE_PROJECT
					lSuccess = THIS.ProjectSearch(cPattern, cProject)
				OTHERWISE
					lSuccess = .F.
				ENDCASE
			ENDIF
			
			SELECT (m.nSelect)
		ENDIF
		
		RETURN m.lSuccess
	ENDFUNC

	FUNCTION SetChecked(cUniqueID, lChecked)
		IF PCOUNT() < 2
			lChecked = .T.
		ENDIF
		IF USED("FoxRefCursor") AND SEEK(cUniqueID, "FoxRefCursor", "UniqueID")
			IF !ISNULL(FoxRefCursor.Checked)
				REPLACE Checked WITH lChecked IN FoxRefCursor
			ENDIF
		ENDIF
	ENDFUNC


	* -- Show the Results form
	FUNCTION ShowResults()
	** Not required
	ENDFUNC

	* goto a definition
	FUNCTION GotoDefinition(cUniqueID)
		LOCAL cFilename
		LOCAL nSelect
		LOCAL oException

		IF VARTYPE(m.cUniqueID) <> 'C' OR EMPTY(m.cUniqueID)
			RETURN .F.
		ENDIF

		IF USED("FoxDefCursor") AND SEEK(m.cUniqueID, "FoxDefCursor", "UniqueID")
			IF INLIST(FoxDefCursor.DefType, DEFTYPE_INCLUDEFILE, DEFTYPE_SETCLASSPROC)
				m.cFilename = FoxDefCursor.Symbol
				
				* if the file doesn't exist (we may not have a full path
				* becase it's from a Window reference), then see if we
				* can find a similar reference in the REFFILE that matches
				* our project
				IF !FILE(m.cFilename) AND USED("ProjectFilesCursor")
					m.nSelect = SELECT()
					SELECT ProjectFilesCursor
					m.cFilename = PADR(LOWER(m.cFilename), LENC(ProjectFilesCursor.Filename))
					LOCATE FOR Filename == m.cFilename
					IF FOUND()
						m.cFilename = ADDBS(ProjectFilesCursor.Folder) + RTRIM(ProjectFilesCursor.Filename)
					ENDIF
					SELECT (m.nSelect)
				ENDIF
				m.cFilename = ALLTRIM(m.cFilename)
				TRY
					m.cFilename = LOCFILE(m.cFilename, JUSTEXT(m.cFilename))
					EDITSOURCE(m.cFilename)
				CATCH TO oException
					MESSAGEBOX(oException.Message, MB_ICONEXCLAMATION, APPNAME_LOC)
				ENDTRY
			ELSE
				THIS.OpenSource( ;
				  FoxDefCursor.FileID, ;
				  FoxDefCursor.LineNo, ;
				  FoxDefCursor.ProcLineNo, ;
				  FoxDefCursor.ClassName, ;
				  FoxDefCursor.ProcName, ;
				  '', ;
				  FoxDefCursor.Symbol, ;
				  0 ;
				 )
			ENDIF
		ENDIF
	ENDFUNC

	* goto a specific reference
	FUNCTION GotoReference(cUniqueID)
		IF VARTYPE(m.cUniqueID) <> 'C' OR EMPTY(m.cUniqueID)
			RETURN .F.
		ENDIF

		IF USED("FoxRefCursor") AND SEEK(m.cUniqueID, "FoxRefCursor", "UniqueID")
			THIS.OpenSource(;
			  FoxRefCursor.FileID, ;
			  FoxRefCursor.LineNo, ;
			  FoxRefCursor.ProcLineNo, ;
			  FoxRefCursor.ClassName, ;
			  FoxRefCursor.ProcName, ;
			  FoxRefCursor.UpdField, ;
			  FoxRefCursor.Symbol ;
			 )

		ENDIF
	
	ENDFUNC


	* open source file for editing
	FUNCTION OpenSource(cFileID, nLineNo, nProcLineNo, cClassName, cProcName, cUpdField, cSymbol, nColPos)
		LOCAL cFileType
		LOCAL nSelect
		LOCAL cAlias
		LOCAL lLibraryOpen
		LOCAL nStartPos
		LOCAL nRetCode
		LOCAL oException
		LOCAL nOpenError
		LOCAL ARRAY aEdEnv[25]

		IF VARTYPE(m.cSymbol) <> 'C'
			m.cSymbol = ''
		ENDIF
		IF VARTYPE(m.nColPos) <> 'N'
			m.nColPos = 0
		ENDIF

		m.nSelect = SELECT()
		IF USED("FileCursor") AND SEEK(cFileID, "FileCursor", "UniqueID")
			IF FileCursor.UniqueID = "WINDOW" AND THIS.WindowHandle >= 0
				* special case for jumping to a specific line # in the open window
				IF ATCC("FOXTOOLS.FLL", SET("LIBRARY")) == 0
					m.lLibraryOpen = .F.
					m.cFoxtoolsLibrary = SYS(2004) + "FOXTOOLS.FLL"
					IF !FILE(m.cFoxtoolsLibrary)
						RETURN .F.
					ENDIF
					SET LIBRARY TO (m.cFoxToolsLibrary) ADDITIVE
				ELSE
					m.lLibraryOpen = .T.
				ENDIF

				* Check environment of window
				* get the length of the window
				m.nRetCode = _edgetenv(THIS.WindowHandle, @aEdEnv)

				IF m.nRetCode == 1 && AND (aEdEnv[EDENV_LENGTH] == 0)
					IF aEdEnv[EDENV_LENGTH] > 0
						m.nStartPos = _EdGetLPos(THIS.WindowHandle, m.nLineNo - 1)
						_EdSelect(THIS.WindowHandle, m.nStartPos, m.nStartPos)
						_EdSToPos(THIS.WindowHandle, m.nStartPos, .F.)
					ENDIF
					_wselect(THIS.WindowHandle)

					THIS.HighlightSymbol(m.cSymbol, m.nLineNo, m.nColPos)
				ENDIF

				IF !m.lLibraryOpen AND ATCC(m.cFoxToolsLibrary, SET("LIBRARY")) > 0
					RELEASE LIBRARY (m.cFoxToolsLibrary)
				ENDIF
			ELSE
				* open the file for editing
				m.cFilename  = ADDBS(RTRIM(FileCursor.Folder)) + RTRIM(FileCursor.FileName)
				m.cFileType  = UPPER(JUSTEXT(m.cFilename))
				
				* make sure the file exists
				IF FILE(m.cFilename)
					IF VARTYPE(m.cUpdField) <> 'C'
						m.cUpdField = ''
					ELSE
						m.cUpdField  = UPPER(m.cUpdField)
					ENDIF

					IF m.cFileType == "CDX"
						m.cFilename = FORCEEXT(m.cFilename, "DBF")
						m.cFileType = "DBF"
					ENDIF

					m.nOpenError = 0
					TRY
						DO CASE
						CASE m.cFileType == "SCX"
							m.nOpenError = EDITSOURCE(m.cFileName, MAX(m.nProcLineNo, 1), m.cClassName, m.cProcName)
							DO CASE
							CASE m.nOpenError == 925 && couldn't find methodname
								THIS.HighlightObject(m.cClassName, m.cProcName)

							CASE m.nOpenError == 0
								THIS.HighlightSymbol(m.cSymbol, MAX(m.nProcLineNo, 1), m.nColPos)
							ENDCASE

						CASE m.cFileType == "VCX"
							m.nOpenError = EDITSOURCE(m.cFileName, MAX(m.nProcLineNo, 1), m.cClassName, m.cProcName)

							DO CASE
							CASE m.nOpenError == 925 && couldn't find methodname
								THIS.HighlightObject(m.cClassName, m.cProcName)
								
							CASE m.nOpenError == 0
								THIS.HighlightSymbol(m.cSymbol, MAX(m.nProcLineNo, 1), m.nColPos)
							ENDCASE

						CASE m.cFileType == "DBC"
							DO CASE
							CASE m.cUpdField == "STOREDPROCEDURE"
								m.nOpenError = EDITSOURCE(m.cFileName, m.nLineNo)
								IF m.nOpenError == 0
									THIS.HighlightSymbol(m.cSymbol, m.nLineNo, m.nColPos)
								ENDIF

							OTHERWISE
								MODIFY DATABASE (m.cFilename)
							ENDCASE
							

						CASE m.cFileType == "DBF"
							m.cAlias = JUSTSTEM(m.cFilename)
							IF USED(m.cAlias)
								SELECT (m.cAlias)
							ELSE
								SELECT 0
								TRY
									USE (m.cFilename) EXCLUSIVE
								CATCH
								ENDTRY
					
								IF !(UPPER(ALIAS()) == UPPER(m.cAlias))
									TRY
										USE (m.cFilename) SHARED AGAIN
									CATCH
									ENDTRY
								ENDIF
							ENDIF
							IF UPPER(ALIAS()) == UPPER(m.cAlias)
								MODIFY STRUCTURE
								
								USE IN (m.cAlias)
							ENDIF

						CASE m.cFileType == "PRG" OR ;
						     m.cFileType == "MPR" OR ;
						     m.cFileType == "QPR" OR ;
						     m.cFileType == "SPR" OR ;
						     m.cFileType == "H" OR ;
						     m.cFileType == "FRX" OR ;
						     m.cFileType == "LBX" OR ;
						     m.cFileType == "MNX" OR ;
						     m.cFileType == "TXT" OR ;
						     m.cFileType == "LOG" OR ;
						     m.cFileType == "ME" 
							m.nOpenError = EDITSOURCE(m.cFileName, m.nLineNo)
							IF m.nOpenError == 0
								THIS.HighlightSymbol(m.cSymbol, m.nLineNo, m.nColPos)
							ENDIF
						

						OTHERWISE
							THIS.ShellTo(m.cFilename)
						ENDCASE
					CATCH TO oException
						MESSAGEBOX(oException.Message, MB_ICONEXCLAMATION, APPNAME_LOC)					
					ENDTRY
				ELSE
					MessageBox(ERROR_FILENOTFOUND_LOC + ':' + CHR(10) + CHR(10) + m.cFilename, MB_ICONEXCLAMATION, APPNAME_LOC)
				ENDIF
				
				IF m.nOpenError > 0 AND !INLIST(m.nOpenError, 901, 925)
					MessageBox(ERROR_OPENFILE_LOC, MB_ICONEXCLAMATION, APPNAME_LOC)
				ENDIF

			
			ENDIF
		ENDIF

		SELECT (m.nSelect)
	ENDFUNC


	* -- Set focus to the control where the reference was found
	* -- This is for when our reference is in a property of an object
	* -- and not in actual code
	PROCEDURE HighlightObject(cParentName, cObjName)
		LOCAL oObjRef
		LOCAL oParentRef
		LOCAL ARRAY aObjRef[1]

		* see if object specified exists and highlight it
		IF ASELOBJ(aObjRef, 1) == 1
			m.oParentRef = aObjRef[1]
			DO WHILE VARTYPE(m.oParentRef) == 'O' AND !(m.oParentRef.Name == m.cParentName)
				m.oParentRef = m.oParentRef.Parent
			ENDDO


			IF VARTYPE(m.oParentRef) == 'O'
				m.oObjRef = m.oParentRef.&cObjName
				IF VARTYPE(m.oObjRef) == 'O'
					IF PEMSTATUS(m.oObjRef, "SetFocus", 5)
						TRY
							m.oObjRef.SetFocus()
						CATCH
							* just in case, don't display an error because it's not critical
						ENDTRY
					ELSE
						* the object doesn't have a SetFocus method (such as Optiongroup), so at
						* least select the container (for example, select the Page the Optiongroup is on)
						TRY
							m.oObjRef.Parent.SetFocus()
						CATCH
							* just in case, don't display an error because it's not critical
						ENDTRY
					ENDIF
				ENDIF
			ENDIF
		ENDIF
	ENDPROC
	
	

	* given a filename, return the FileID from RefFile
	FUNCTION GetFileID(cFilename)
		LOCAL cFName
		LOCAL cFolder
		LOCAL nSelect
		LOCAL cFileID
		LOCAL ARRAY aFileList[1]
		
		m.nSelect = SELECT()

		m.cFilename = LOWER(m.cFilename)
		cFName  = PADR(LOWER(JUSTFNAME(m.cFilename)), 100)
		* cFolder = PADR(LOWER(JUSTPATH(m.cFilename)), 240)
		cFolder = LOWER(JUSTPATH(m.cFilename))

		SELECT UniqueID ;
		 FROM (THIS.FileTable) FileTable ;
		 WHERE ;
		   FileTable.Filename == cFName AND ;
		   FileTable.Folder == cFolder AND ;
		   FileTable.FileAction <> FILEACTION_INACTIVE ;
		 INTO ARRAY aFileList
		IF _TALLY > 0
			m.cFileID = aFileList[1]
		ELSE
			m.cFileID = ''
		ENDIF
		
		SELECT (m.nSelect)
		
		RETURN m.cFileID
	ENDFUNC

	* find all matching definitions and go directly to it
	* or display ambiguous references
	FUNCTION GotoSymbol(cGotoSymbol, cFilename, cClassName, cProcName, nLineNo, lLocalOnly)
		LOCAL nSelect
		LOCAL cSymbol
 		LOCAL nFunctionCnt
		LOCAL i
		LOCAL nPos
		LOCAL cHelpTopic
		LOCAL cWhere
		LOCAL cWhere2
		LOCAL cFName
		LOCAL cFolder
		LOCAL lFindProperty
		LOCAL lFindProcedure
		LOCAL nIndex
		LOCAL lSuccess
		LOCAL cHelpSymbol
		LOCAL ARRAY aCommandList[1]
		LOCAL ARRAY aFunctionList[1]
		LOCAL ARRAY aBaseClassList[1]

		IF VARTYPE(m.cGotoSymbol) <> 'C' OR EMPTY(m.cGotoSymbol)
			RETURN .F.
		ENDIF

		* search only for locals within current window, and
		* return if there is no open window
		IF m.lLocalOnly AND THIS.WindowHandle < 0
			RETURN .F.
		ENDIF
		
		m.cSymbol = ALLTRIM(UPPER(m.cGotoSymbol))

		m.lFindProperty = LEFTC(m.cSymbol, 5) == "THIS." OR LEFTC(m.cSymbol, 9) == "THISFORM."
		m.lFindProcedure = .F.

		* if this is a function (as determined by having a '(' in the symbol name),
		* then strip off everything after the paren
		m.nPos = RATC('(', m.cSymbol)
		IF m.nPos > 1
			m.cSymbol = LEFTC(m.cSymbol, m.nPos - 1)
			m.lFindProcedure = .T.
		ENDIF


		* remove the "M.", "THIS." or "THISFORM." from the symbol all the
		* way up to the last word after the '.'
		* For example:
		*   "THIS.Parent.cmdButton.AutoSize" = "AutoSize"
		m.nPos = RATC('.', m.cSymbol)
		IF m.nPos > 0 AND m.nPos < LENC(m.cSymbol)
			m.cSymbol = SUBSTRC(m.cSymbol, m.nPos + 1)
		ENDIF

		* remove ampersand character from beginning if it exists		
		m.nPos = RATC('&', m.cSymbol)
		IF m.nPos > 0 AND m.nPos < LENC(m.cSymbol)
			m.cSymbol = SUBSTRC(m.cSymbol, m.nPos + 1)
		ENDIF
		
		
		IF EMPTY(m.cSymbol)
			RETURN .F.
		ENDIF

		m.lSuccess = .F.

		IF VARTYPE(m.cFilename) == 'C'
			cFName  = PADR(LOWER(JUSTFNAME(m.cFilename)), 100)
			cFolder = LOWER(JUSTPATH(m.cFilename))
		ELSE
			cFName  = ''
			cFolder = ''
		ENDIF
		
		IF VARTYPE(m.cProcName) <> 'C'
			m.cProcName = ''
		ENDIF



		m.nSelect = SELECT()
		m.cWhere = [DefTable.Symbol == cSymbol AND !DefTable.Inactive]
		m.nCnt = 0

		IF m.lFindProcedure
			* search only for procedure names
			m.cWhere = m.cWhere + [ AND (DefTable.DefType == "] + DEFTYPE_PROCEDURE + [")]
		ELSE
			* search only for properties and procedures
			IF m.lFindProperty
				m.cWhere = m.cWhere + [ AND (DefTable.DefType == "] + DEFTYPE_PROPERTY + [" OR DefTable.DefType == "] + DEFTYPE_PROCEDURE + [")]
			ENDIF
		ENDIF
		m.cWhere2 = m.cWhere

		IF THIS.WindowHandle >= 0
			* see if we're positioned on a #include, SET PROCEDURE TO line
			IF VARTYPE(m.nLineNo) == 'N' AND m.nLineNo > 0
				SELECT ;
				  DefTable.UniqueID, ;
				  DefTable.FileID, ;
				  DefTable.DefType, ;
				  DefTable.ProcName, ;
				  DefTable.ClassName, ;
				  DefTable.ProcLineNo, ;
				  DefTable.LineNo, ;
				  DefTable.Abstract ;
				FROM (THIS.DefTable) DefTable ;
				 WHERE ;
				  DefTable.FileID == "WINDOW" AND ;
				  (DefTable.DefType == DEFTYPE_INCLUDEFILE OR DefTable.DefType == DEFTYPE_SETCLASSPROC) AND ;
				  DefTable.LineNo == m.nLineNo AND ;
				  !DefTable.Inactive ;
				 INTO CURSOR DefinitionCursor
				nCnt = _TALLY
			ENDIF

			IF nCnt == 0 
				IF m.lLocalOnly AND EMPTY(m.cProcName) AND VARTYPE(THIS.oWindowEngine) == 'O'
					m.cProcName = THIS.oWindowEngine.GetProcedure(m.nLineNo)
					
					cWhere = cWhere + [ AND DefTable.ProcName == m.cProcName]
				ENDIF

				* always search for the LOCAL or PRIVATE or PARAMETER first
				SELECT ;
				  DefTable.UniqueID, ;
				  DefTable.FileID, ;
				  DefTable.DefType, ;
				  DefTable.ProcName, ;
				  DefTable.ClassName, ;
				  DefTable.ProcLineNo, ;
				  DefTable.LineNo, ;
				  DefTable.Abstract, ;
				  '' AS Folder, ;
				  '' AS Filename ;
				 FROM (THIS.DefTable) DefTable INNER JOIN (THIS.FileTable) FileTable ON DefTable.FileID == FileTable.UniqueID ;
				 WHERE DefTable.FileID = "WINDOW" AND ;
				   &cWhere AND ;
				  INLIST(DefTable.DefType, DEFTYPE_PARAMETER, DEFTYPE_LOCAL, DEFTYPE_PRIVATE, DEFTYPE_PUBLIC, DEFTYPE_PROCEDURE, DEFTYPE_CLASS, DEFTYPE_DEFINE, DEFTYPE_PROPERTY) ;
				 INTO CURSOR DefinitionCursor
				nCnt = _TALLY
			ENDIF
				
			* now search in the current file
			IF nCnt == 0
				SELECT ;
				  DefTable.UniqueID, ;
				  DefTable.FileID, ;
				  DefTable.DefType, ;
				  DefTable.ProcName, ;
				  DefTable.ClassName, ;
				  DefTable.ProcLineNo, ;
				  DefTable.LineNo, ;
				  DefTable.Abstract, ;
				  '' AS Folder, ;
				  '' AS Filename ;
				 FROM (THIS.DefTable) DefTable INNER JOIN (THIS.FileTable) FileTable ON DefTable.FileID == FileTable.UniqueID ;
				 WHERE DefTable.FileID = "WINDOW" AND ;
				   &cWhere2 AND ;
				  INLIST(DefTable.DefType, DEFTYPE_PRIVATE, DEFTYPE_PUBLIC, DEFTYPE_PROCEDURE, DEFTYPE_CLASS, DEFTYPE_DEFINE, DEFTYPE_PROPERTY) ;
				 INTO CURSOR DefinitionCursor
				nCnt = _TALLY
			ENDIF
		ELSE
			* always search for the LOCAL or PRIVATE or PARAMETER first
			SELECT ;
			  DefTable.UniqueID, ;
			  DefTable.FileID, ;
			  DefTable.DefType, ;
			  DefTable.ProcName, ;
			  DefTable.ClassName, ;
			  DefTable.ProcLineNo, ;
			  DefTable.LineNo, ;
			  DefTable.Abstract, ;
			  FileTable.Folder, ;
			  FileTable.Filename ;
			FROM (THIS.DefTable) DefTable INNER JOIN (THIS.FileTable) FileTable ON DefTable.FileID == FileTable.UniqueID ;
			 WHERE FileTable.Filename == cFName AND FileTable.Folder == cFolder AND ;
			  &cWhere AND ;
			  INLIST(DefTable.DefType, DEFTYPE_PARAMETER, DEFTYPE_LOCAL, DEFTYPE_PRIVATE, DEFTYPE_PUBLIC, DEFTYPE_PROCEDURE, DEFTYPE_CLASS, DEFTYPE_DEFINE, DEFTYPE_PROPERTY) ;
			 INTO CURSOR DefinitionCursor
			nCnt = _TALLY
		ENDIF


		IF nCnt == 0 AND !m.lLocalOnly
			* now find all appropriate definitions
			IF USED("ProjectFilesCursor") AND RECCOUNT("ProjectFilesCursor") > 0
				* if we have a list of files in this project, then we 
				* only want to concern ourselves with symbols that
				* appear in those files
				SELECT ;
				  DefTable.UniqueID, ;
				  DefTable.FileID, ;
				  DefTable.DefType, ;
				  DefTable.ProcName, ;
				  DefTable.ClassName, ;
				  DefTable.ProcLineNo, ;
				  DefTable.LineNo, ;
				  DefTable.Abstract, ;
				  FileTable.Folder, ;
				  FileTable.Filename ;
				 FROM (THIS.DefTable) DefTable INNER JOIN (THIS.FileTable) FileTable ON DefTable.FileID == FileTable.UniqueID ;
				                               INNER JOIN ProjectFilesCursor ON FileTable.Filename == ProjectFilesCursor.Filename AND FileTable.Folder == ProjectFilesCursor.Folder ;
				 WHERE &cWhere AND ;
				  INLIST(DefTable.DefType, DEFTYPE_PUBLIC, DEFTYPE_PROCEDURE, DEFTYPE_CLASS, DEFTYPE_DEFINE) ;
				 INTO CURSOR DefinitionCursor
			ELSE
				SELECT ;
				  DefTable.UniqueID, ;
				  DefTable.FileID, ;
				  DefTable.DefType, ;
				  DefTable.ProcName, ;
				  DefTable.ClassName, ;
				  DefTable.ProcLineNo, ;
				  DefTable.LineNo, ;
				  DefTable.Abstract, ;
				  FileTable.Folder, ;
				  FileTable.Filename ;
				FROM (THIS.DefTable) DefTable INNER JOIN (THIS.FileTable) FileTable ON DefTable.FileID == FileTable.UniqueID ;
				 WHERE &cWhere AND ;
				  INLIST(DefTable.DefType, DEFTYPE_PUBLIC, DEFTYPE_PROCEDURE, DEFTYPE_CLASS, DEFTYPE_DEFINE) ;
				 INTO CURSOR DefinitionCursor
			ENDIF
			nCnt = _TALLY
		ENDIF

* Moved higher so opens #include files faster
*!*			IF nCnt == 0 AND VARTYPE(m.nLineNo) == 'N' AND m.nLineNo > 0 AND !m.lLocalOnly
*!*				* see if we're positioned on a #include, SET PROCEDURE TO line
*!*				SELECT ;
*!*				  DefTable.UniqueID, ;
*!*				  DefTable.FileID, ;
*!*				  DefTable.DefType, ;
*!*				  DefTable.ProcName, ;
*!*				  DefTable.ClassName, ;
*!*				  DefTable.ProcLineNo, ;
*!*				  DefTable.LineNo, ;
*!*				  DefTable.Abstract ;
*!*				FROM (THIS.DefTable) DefTable ;
*!*				 WHERE ;
*!*				  DefTable.FileID == "WINDOW" AND ;
*!*				  (DefTable.DefType == DEFTYPE_INCLUDEFILE OR DefTable.DefType == DEFTYPE_SETCLASSPROC) AND ;
*!*				  DefTable.LineNo == m.nLineNo AND ;
*!*				  !DefTable.Inactive ;
*!*				 INTO CURSOR DefinitionCursor
*!*				nCnt = _TALLY
*!*			ENDIF
						
		DO CASE
		CASE nCnt == 0  && no matches found
			IF !m.lLocalOnly
				* if this is an internal VFP command/function, then
				* display the VFP Help on that symbol
				cHelpTopic = ''

				* look for base classes
				=ALANGUAGE(aBaseClassList, 3)

				nIndex = ASCAN(aBaseClassList, GETWORDNUM(cGotoSymbol, 1), -1, -1, 1, 7)
				IF nIndex > 0
					cHelpTopic = aBaseClassList[nIndex]
					HELP &cHelpTopic
				ELSE
					* create array of all internal VFP commands
					m.cHelpSymbol = UPPER(GETWORDNUM(cGotoSymbol, 1))
					=ALANGUAGE(aCommandList, 1)
					IF ASCAN(aCommandList, m.cHelpSymbol, -1, -1, 1, 14) > 0
						cHelpTopic = UPPER(cGotoSymbol)
					ELSE
						* check if there's a matching function
						* don't use ASCAN because we have to check for abbreviated match
						nFunctionCnt = ALANGUAGE(aFunctionList, 2)
						FOR i = 1 TO nFunctionCnt
							IF 'M' $ aFunctionList[i, 2]
								* 'M' in 2nd parameter means we need an exact match
								IF aFunctionList[i, 1] == m.cHelpSymbol
									cHelpTopic = aFunctionList[i, 1] + "()"
								ENDIF
							ELSE
								* no 'M' in 2nd parameter returned by ALANGUAGE, so only need to match up to 4 characters
								IF aFunctionList[i, 1] == m.cHelpSymbol OR (LENC(m.cHelpSymbol) < LENC(aFunctionList[i, 1]) AND LENC(m.cHelpSymbol) >= 4 AND aFunctionList[i, 1] = m.cHelpSymbol)
									cHelpTopic = aFunctionList[i, 1] + "()"
								ENDIF
							ENDIF
						ENDFOR
					ENDIF
				ENDIF
				
				
				IF EMPTY(cHelpTopic)
					MESSAGEBOX(NODEFINITION_LOC + CHR(10) + CHR(10) + cGotoSymbol, MB_ICONEXCLAMATION, GOTODEFINITION_LOC)
				ELSE
					cHelpTopic = ALLTRIM(cHelpTopic)
					HELP &cHelpTopic
				ENDIF
			ENDIF

		CASE nCnt == 1
			* only a single match, so go right to it
			THIS.GotoDefinition(DefinitionCursor.UniqueID)
			m.lSuccess = .T.

		OTHERWISE
			* more than one match found, so display a cursor of
			* the available matches.
			*!* DO FORM FoxRefGotoDef WITH THIS, cGotoSymbol
			m.lSuccess = .T.
		ENDCASE

		IF USED("DefinitionCursor")
			USE IN DefinitionCursor
		ENDIF

		SELECT (m.nSelect)
		
		RETURN m.lSuccess
	ENDFUNC


	* Show a progress form while searching
	FUNCTION ProgressInit(cDescription, nMax)
		IF THIS.ShowProgress
			IF VARTYPE(THIS.oProgressForm) <> 'O'
				THIS.UpdateProgress()
			ENDIF
			IF VARTYPE(THIS.oProgressForm) == 'O'
				THIS.oProgressForm.SetMax(m.nMax)
				THIS.oProgressForm.SetDescription(m.cDescription)
			ENDIF
			DOEVENTS
		ENDIF
	ENDFUNC

	FUNCTION UpdateProgress(nValue, cMsg, lFilename)
		IF THIS.ShowProgress
			IF VARTYPE(THIS.oProgressForm) <> 'O'
				THIS.lCancel = .F.
				THIS.oProgressForm = NEWOBJECT("CProgressForm", "FoxRef.vcx")
				THIS.oProgressForm.Show()
			ENDIF
			
			IF m.lFilename AND !EMPTY(JUSTPATH(m.cMsg))
				* truncate filenames so they fit
				m.cMsg = DISPLAYPATH(m.cMsg, 60)
			ENDIF
			
			IF !THIS.oProgressForm.SetProgress(m.nValue, m.cMsg)  && FALSE is returned if Cancel button is pressed
				IF MESSAGEBOX(SEARCH_CANCEL_LOC, MB_ICONQUESTION + MB_YESNO, APPNAME_LOC) == IDYES
					THIS.lCancel = .T.
				ELSE
					THIS.oProgressForm.lCancel = .F.
				ENDIF
			ENDIF
			DOEVENTS
		ENDIF
	ENDFUNC

	FUNCTION CloseProgress()
		IF VARTYPE(THIS.oProgressForm) == 'O'
			THIS.oProgressForm.Release()
		ENDIF
	ENDFUNC


	* Export reference table
	FUNCTION ExportReferences(cExportType, cFilename, cSetID, cFileID, cRefID, lSelectedOnly)
		LOCAL nSelect
		LOCAL cWhere
		LOCAL lSuccess
		LOCAL cXSLFilename
		LOCAL cTempFilename
		LOCAL cHTML
		LOCAL xslDoc
		LOCAL xmlDoc
		LOCAL cSafety
		LOCAL oException

		nSelect = SELECT()
		
		oException = .NULL.

		IF VARTYPE(cExportType) <> 'C' OR EMPTY(cExportType)
			cExportType = EXPORTTYPE_DBF
		ENDIF
		IF VARTYPE(cSetID) <> 'C'
			cSetID = ''
		ENDIF
		IF VARTYPE(cFileID) <> 'C'
			cFileID = ''
		ENDIF
		IF VARTYPE(cRefID) <> 'C'
			cRefID = ''
		ENDIF
		IF VARTYPE(lSelectedOnly) <> 'L'
			lSelectedOnly = .F.
		ENDIF

		* all export types except "Clipboard" require a filename		
		IF cExportType <> EXPORTTYPE_CLIPBOARD
			IF VARTYPE(cFilename) <> 'C'
				RETURN .F.
			ENDIF
			cFilename = FULLPATH(cFilename)
			IF !DIRECTORY(JUSTPATH(cFilename))
				RETURN .F.
			ENDIF
		ENDIF


		cWhere = "RefTable.RefType == [" + REFTYPE_RESULT + "] AND !RefTable.Inactive"
		IF EMPTY(cRefID)
			IF !EMPTY(cSetID)
				cWhere = cWhere + " AND RefTable.SetID == [" + cSetID + "]"
			ENDIF
			IF !EMPTY(cFileID)
				cWhere = cWhere + " AND RefTable.FileID == [" + cFileID + "]"
			ENDIF
		ELSE
			cWhere = cWhere + " AND RefTable.RefID == [" + cRefID + "]"
		ENDIF

		IF lSelectedOnly
			cWhere = cWhere + " AND RefTable.Checked"
		ENDIF

		IF INLIST(cExportType, EXPORTTYPE_TXT, EXPORTTYPE_XLS)
			* memo fields won't be copied so we need to convert
			* to long characters
			SELECT ;
			  PADR(RefTable.Symbol, 254) AS Symbol, ;
			  FileTable.Folder, ;
			  FileTable.Filename, ;
			  PADR(RefTable.ClassName, 254) AS ClassName, ;
			  PADR(RefTable.ProcName, 254) AS ProcName, ;
			  RefTable.ProcLineNo, ;
			  RefTable.LineNo, ;
			  RefTable.ColPos, ;
			  RefTable.MatchLen, ;
			  PADR(StripTabs(RefTable.Abstract), 254) AS Abstract ;
			 FROM (THIS.RefTable) RefTable INNER JOIN (THIS.FileTable) FileTable ON RefTable.FileID == FileTable.UniqueID ;
			 WHERE &cWhere ;
			 INTO CURSOR ExportCursor
		ELSE
			SELECT ;
			  RefTable.Symbol, ;
			  FileTable.Folder, ;
			  FileTable.Filename, ;
			  RefTable.ClassName, ;
			  RefTable.ProcName, ;
			  RefTable.ProcLineNo, ;
			  RefTable.LineNo, ;
			  RefTable.ColPos, ;
			  RefTable.Abstract ;
			 FROM (THIS.RefTable) RefTable INNER JOIN (THIS.FileTable) FileTable ON RefTable.FileID == FileTable.UniqueID ;
			 WHERE &cWhere ;
			 INTO CURSOR ExportCursor
		ENDIF
		SELECT ExportCursor
		
		DO CASE
		CASE cExportType == EXPORTTYPE_DBF
			TRY
				COPY TO (cFilename) ;
				 FIELDS ;
				  Symbol, ;
				  Folder, ;
				  Filename, ;
				  ClassName, ;
				  ProcName, ;
				  ProcLineNo, ;
				  LineNo, ;
				  ColPos, ;
				  Abstract

			CATCH TO oException
			ENDTRY

		CASE cExportType == EXPORTTYPE_TXT
			TRY
				COPY TO (cFilename) ;
				 FIELDS ;
				  Symbol, ;
				  Folder, ;
				  Filename, ;
				  ClassName, ;
				  ProcName, ;
				  ProcLineNo, ;
				  LineNo, ;
				  ColPos, ;
				  Abstract ;
				 DELIMITED
			CATCH TO oException
			ENDTRY

		CASE cExportType == EXPORTTYPE_XML
			IF EMPTY(JUSTEXT(cFilename))
				cFilename = FORCEEXT(cFilename, "xml")
			ENDIF

		
			TRY
				IF THIS.XMLSchema
					CURSORTOXML("ExportCursor", cFilename, IIF(THIS.XMLFormat == XMLFORMAT_ELEMENTS, 1, 2), IIF(THIS.XMLFormat == XMLFORMAT_ELEMENTS, 520, 512), 0, '1')
				ELSE
					CURSORTOXML("ExportCursor", cFilename, IIF(THIS.XMLFormat == XMLFORMAT_ELEMENTS, 1, 2), IIF(THIS.XMLFormat == XMLFORMAT_ELEMENTS, 520, 512))
				ENDIF

			CATCH TO oException
			ENDTRY

		CASE cExportType == EXPORTTYPE_HTML
			cTempFilename = ADDBS(SYS(2023)) + SYS(2015) + ".xml"
			IF EMPTY(JUSTEXT(cFilename))
				cFilename = FORCEEXT(cFilename, "html")
			ENDIF

			cHTML = ''
			lSuccess = .T.
			TRY
				CURSORTOXML("ExportCursor", cTempFilename, 1, 522, 0, '1')
				cXSLFilename = FULLPATH(THIS.XSLTemplate)
				IF !FILE(cXSLFilename)
					cXSLFilename = THIS.FoxRefDirectory + THIS.XSLTemplate
					IF !FILE(cXSLFilename)
						* still not found, so copy internal version
						* to home directory and use that
						cXSLFilename = THIS.FoxRefDirectory + THIS.XSLTemplate
						STRTOFILE(FILETOSTR("foxreftemplate.xsl"), cXSLFilename)
					ENDIF
				ENDIF

				xmldoc = CreateObject("MSXML2.DOMDocument")
				xsldoc = CreateObject("MSXML2.DOMDocument")

				xmldoc.Load(cTempFilename)
				IF xmldoc.parseerror.errorcode == 0
					xsldoc.Load(cXSLFilename)
					cHTML = xmldoc.TransformNode(xsldoc)
					STRTOFILE(cHTML, cFilename)
				ENDIF

			CATCH TO oException
			ENDTRY
	
				
			cSafety = SET("SAFETY")
			SET SAFETY OFF
			ERASE (cTempFilename)
			SET SAFETY &cSafety

			xsldoc = Null
			xmldoc = Null

		CASE cExportType == EXPORTTYPE_XLS
			TRY
				COPY TO (cFilename) ;
				 FIELDS ;
				  Symbol, ;
				  Folder, ;
				  Filename, ;
				  ClassName, ;
				  ProcName, ;
				  ProcLineNo, ;
				  LineNo, ;
				  ColPos, ;
				  Abstract ;
				 TYPE XL5

			CATCH TO oException
			ENDTRY

		CASE cExportType == EXPORTTYPE_CLIPBOARD
			* NOTE: we can't use DataToClip because it only works for DataSession 1
			* _VFP.DataToClip()

			_ClipText = ''
			SCAN ALL
				_ClipText = ;
				  _ClipText + IIF(RECNO() > 1, CHR(13) + CHR(10), '') + ;
				  Symbol + ',' + ;
				  RTRIM(Folder) + ',' + ;
				  RTRIM(Filename) + ',' + ;
				  ClassName + ',' + ;
				  ProcName + ',' + ;
				  TRANSFORM(ProcLineNo) + ',' + ;
				  TRANSFORM(LineNo) + ',' + ;
				  TRANSFORM(ColPos) + ',' + ;
				  StripTabs(Abstract)
			ENDSCAN

		ENDCASE
		
		IF USED("ExportCursor")
			USE IN ExportCursor
		ENDIF

		IF VARTYPE(oException) == 'O'
			MessageBox(oException.Message, MB_ICONEXCLAMATION, APPNAME_LOC)
		ENDIF


		SELECT (nSelect)
		
		RETURN ISNULL(oException)
	ENDFUNC


	* Print a report of found references
	FUNCTION PrintReferences(cReportName, lPreview, cSetID, lSelectedOnly, cSortColumns)
		LOCAL nSelect
		LOCAL cWhere
		LOCAL cFilename
		LOCAL cExt
		LOCAL oReportAddIn
		LOCAL oRpt
		LOCAL cExecute
		LOCAL nRecordCnt
		

		nSelect = SELECT()
		SELECT 0
		
		nRecordCnt = -1


		IF VARTYPE(lPreview) <> 'L'
			lPreview = .F.
		ENDIF
		IF VARTYPE(lSelectedOnly) <> 'L'
			lSelectedOnly = .F.
		ENDIF
		IF VARTYPE(cSetID) <> 'C'
			cSetID = ''
		ENDIF

		IF VARTYPE(cSortColumns) <> 'C' OR EMPTY(cSortColumns)
			cSortColumns = "FileTable.Folder, FileTable.Filename, RefTable.LineNo"
		ENDIF

		TRY
			oReportAddIn = THIS.oReportCollection.Item(m.cReportName)
		CATCH
			oReportAddIn = .NULL.
		ENDTRY
		
		IF !ISNULL(oReportAddIn)
			cWhere = "RefTable.RefType == [" + REFTYPE_RESULT + "] AND !RefTable.Inactive"
			IF !EMPTY(cSetID)
				cWhere = cWhere + " AND RefTable.SetID == [" + cSetID + "]"
			ENDIF
			IF lSelectedOnly
				cWhere = cWhere + " AND RefTable.Checked"
			ENDIF

			SELECT ;
			  RefTable.RefID, ;
			  RefTable.SetID, ;
			  RefTable.Symbol, ;
			  RefTable.ClassName, ;
			  RefTable.ProcName, ;
			  RefTable.ProcLineNo, ;
			  RefTable.LineNo, ;
			  RefTable.ColPos, ;
			  RefTable.MatchLen, ;
			  RefTable.Abstract, ;
			  RefTable.TimeStamp, ;
			  FileTable.Folder, ;
			  FileTable.Filename, ;
			  PADR(FileTable.Folder, 240) AS SortFolder, ;
			  LOWER(PADR(JUSTEXT(FileTable.Filename), 3)) AS SortFileType, ;
			  PADR(RefTable.ClassName, 100) AS ClassNameSort, ;
			  PADR(RefTable.ProcName, 100) AS ProcNameSort ;
			 FROM (THIS.RefTable) RefTable INNER JOIN (THIS.FileTable) FileTable ON RefTable.FileID == FileTable.UniqueID ;
			 WHERE &cWhere ;
			 ORDER BY &cSortColumns ;
			 INTO CURSOR RptCursor
			m.nRecordCnt = _TALLY

			IF m.nRecordCnt > 0
				WITH oReportAddIn
					TRY
						DO CASE
						CASE !EMPTY(.RptFilename)
							m.cExt = UPPER(JUSTEXT(.RptFilename))
							IF EMPTY(m.cExt)
								m.cExt = "FRX"
							ENDIF

							DO CASE
							CASE m.cExt == "PRG" OR m.cExt == "FXP"
								DO (.RptFilename) WITH m.lPreview, THIS

							CASE m.cExt == "SCX"
								DO FORM (.RptFilename) WITH m.lPreview, THIS
							
							CASE m.cExt == "FRX"
								IF m.lPreview
									REPORT FORM (.RptFilename) PREVIEW
								ELSE
									REPORT FORM (.RptFilename) NOCONSOLE TO PRINTER PROMPT
								ENDIF
							ENDCASE

						CASE !EMPTY(.RptClassName) AND !EMPTY(.RptMethod)
							cExecute = [oRpt.] + .RptMethod + [(] + TRANSFORM(m.lPreview) + [, THIS"]
							oRpt = NEWOBJECT(.RptClassName, .RptClassLibrary)
							&cExecute
						OTHERWISE
							nRecordCnt = -1
						ENDCASE
					CATCH TO oException
						MessageBox(oException.Message, MB_ICONSTOP, APPNAME_LOC)
					ENDTRY
				ENDWITH
			ENDIF

			IF USED("RptCursor")
				USE IN RptCursor
			ENDIF

		ENDIF
		
		SELECT (nSelect)
		
		RETURN nRecordCnt
	ENDFUNC

	* Clear a specified Result Set or all results
	* from the Reference table
	FUNCTION ClearResults(cSetID, cFileID)
		LOCAL nSelect
		LOCAL cAlias
		
		m.nSelect = SELECT()

		DO CASE
		CASE VARTYPE(m.cFileID) == 'C' AND !EMPTY(m.cFileID)		
			* Clear specified file
			DELETE FROM (THIS.RefTable) WHERE SetID == m.cSetID AND FileID == m.cFileID

		CASE VARTYPE(m.cSetID) == 'C' AND !EMPTY(m.cSetID)
			* Clear specified result set
			DELETE FROM (THIS.RefTable) WHERE SetID == m.cSetID

		OTHERWISE
			* Clear all results
			DELETE FROM (THIS.RefTable) ;
			 WHERE ;
			  RefType == REFTYPE_RESULT OR ;
			  RefType == REFTYPE_SEARCH OR ;
			  RefType == REFTYPE_NOMATCH OR ;
			  RefType == REFTYPE_ERROR OR ;
			  RefType == REFTYPE_LOG
		ENDCASE

		
		* if we can get this table open exclusive, then 
		* we should pack it
		IF THIS.OpenRefTable(THIS.RefTable, .T.)
			SELECT FoxRefCursor
			TRY
				PACK IN FoxRefCursor
			CATCH
				* no big deal that we can't pack the table -- just ignore the error
			ENDTRY
		ENDIF

		
		* open it again shared
		THIS.OpenRefTable(THIS.RefTable)


		SELECT (m.nSelect)
		
		RETURN .T.
	ENDFUNC


	
	* Save preferences to FoxPro Resource file
	FUNCTION SavePrefs()
		LOCAL nSelect
		LOCAL lSuccess
		LOCAL nMemoWidth
		LOCAL nCnt
		LOCAL cData
		LOCAL i
		LOCAL oOptionCollection
		LOCAL ARRAY aFileList[1]
		LOCAL ARRAY FOXREF_OPTIONS[1]
		LOCAL ARRAY FOXREF_LOOKFOR_MRU[10]
		LOCAL ARRAY FOXREF_FOLDER_MRU[10]
		LOCAL ARRAY FOXREF_FILETYPES_MRU[10]

		IF !(SET("RESOURCE") == "ON")
			RETURN .F.
		ENDIF

		=ACOPY(THIS.aLookForMRU, FOXREF_LOOKFOR_MRU)
		=ACOPY(THIS.aReplaceMRU, FOXREF_REPLACE_MRU)
		=ACOPY(THIS.aFolderMRU, FOXREF_FOLDER_MRU)
		=ACOPY(THIS.aFileTypesMRU, FOXREF_FILETYPES_MRU)

		oOptionCollection = CREATEOBJECT("Collection")
		* Add any properties you want to save to
		* the resource file to this collection
		WITH oOptionCollection
			.Add(THIS.Comments, "Comments")
			.Add(THIS.MatchCase, "MatchCase")
			.Add(THIS.WholeWordsOnly, "WholeWordsOnly")
			.Add(THIS.Wildcards, "Wildcards")
			.Add(THIS.ProjectHomeDir, "ProjectHomeDir")
			.Add(THIS.SubFolders, "SubFolders")
			.Add(THIS.OverwritePrior, "OverwritePrior")
			.Add(THIS.FileTypes, "FileTypes")
			.Add(THIS.IncludeDefTable, "IncludeDefTable")
			.Add(THIS.CodeOnly, "CodeOnly")
			.Add(THIS.FormProperties, "FormProperties")
			.Add(THIS.AutoProjectHomeDir, "AutoProjectHomeDir")
			.Add(THIS.ConfirmReplace, "ConfirmReplace")
			.Add(THIS.BackupOnReplace, "BackupOnReplace")
			.Add(THIS.DisplayReplaceLog, "DisplayReplaceLog")
			.Add(THIS.PreserveCase, "PreserveCase")
			.Add(THIS.BackupStyle, "BackupStyle")
			.Add(THIS.ShowRefsPerLine, "ShowRefsPerLine")
			.Add(THIS.ShowFileTypeHistory, "ShowFileTypeHistory")
			.Add(THIS.ShowDistinctMethodLine, "ShowDistinctMethodLine")
			.Add(THIS.SortMostRecentFirst, "SortMostRecentFirst")
			.Add(THIS.FontString, "FontString")
			.Add(THIS.FoxRefDirectory, "FoxRefDirectory")
		ENDWITH

		DIMENSION FOXREF_OPTIONS[MAX(oOptionCollection.Count, 1), 2]
		FOR i = 1 TO oOptionCollection.Count
			FOXREF_OPTIONS[m.i, 1] = oOptionCollection.GetKey(m.i)
			FOXREF_OPTIONS[m.i, 2] = oOptionCollection.Item(m.i)
		ENDFOR


		nSelect = SELECT()
		
		lSuccess = .F.

  		* make sure Resource file exists and is not read-only
  		TRY
	  		nCnt = ADIR(aFileList, SYS(2005))
	  	CATCH
	  		nCnt = 0
	  	ENDTRY
	  		
		IF nCnt > 0 AND ATCC('R', aFileList[1, 5]) == 0
			IF !USED("FoxResource")
				USE (SYS(2005)) IN 0 SHARED AGAIN ALIAS FoxResource
			ENDIF
			IF USED("FoxResource") AND !ISREADONLY("FoxResource")
				nMemoWidth = SET('MEMOWIDTH')
				SET MEMOWIDTH TO 255

				SELECT FoxResource
				LOCATE FOR UPPER(ALLTRIM(type)) == "PREFW" AND UPPER(ALLTRIM(id)) == RESOURCE_ID AND EMPTY(name)
				IF !FOUND()
					APPEND BLANK IN FoxResource
					REPLACE ; 
					  Type WITH "PREFW", ;
					  ID WITH RESOURCE_ID, ;
					  ReadOnly WITH .F. ;
					 IN FoxResource
				ENDIF

				IF !FoxResource.ReadOnly
					SAVE TO MEMO Data ALL LIKE FOXREF_*

					REPLACE ;
					  Updated WITH DATE(), ;
					  ckval WITH VAL(SYS(2007, FoxResource.Data)) ;
					 IN FoxResource

					lSuccess = .T.
				ENDIF
				SET MEMOWIDTH TO (nMemoWidth)
			
				USE IN FoxResource
			ENDIF
		ENDIF

		SELECT (nSelect)
		
		RETURN lSuccess
	ENDFUNC


	* retrieve preferences from the FoxPro Resource file
	FUNCTION RestorePrefs()
		LOCAL nSelect
		LOCAL lSuccess
		LOCAL nMemoWidth
		LOCAL i
		LOCAL nCnt

		LOCAL ARRAY FOXREF_LOOKFOR_MRU[10]
		LOCAL ARRAY FOXREF_REPLACE_MRU[10]
		LOCAL ARRAY FOXREF_FOLDER_MRU[10]
		LOCAL ARRAY FOXREF_FILETYPES_MRU[10]
		LOCAL ARRAY FOXREF_OPTIONS[1]

		IF !(SET("RESOURCE") == "ON")
			RETURN .F.
		ENDIF


		m.nSelect = SELECT()
		
		m.lSuccess = .F.

		IF FILE(SYS(2005))    && resource file not found.
			USE (SYS(2005)) IN 0 SHARED AGAIN ALIAS FoxResource
			IF USED("FoxResource")
				m.nMemoWidth = SET('MEMOWIDTH')
				SET MEMOWIDTH TO 255

				SELECT FoxResource
				LOCATE FOR UPPER(ALLTRIM(type)) == "PREFW" AND ;
		   			UPPER(ALLTRIM(id)) == RESOURCE_ID AND ;
		   			EMPTY(name) AND ;
		   			!DELETED()

				IF FOUND() AND !EMPTY(Data) AND ckval == VAL(SYS(2007, Data))
					RESTORE FROM MEMO Data ADDITIVE

					IF TYPE("FOXREF_LOOKFOR_MRU") == 'C'
						=ACOPY(FOXREF_LOOKFOR_MRU, THIS.aLookForMRU)
					ENDIF
					IF TYPE("FOXREF_REPLACE_MRU") == 'C'
						=ACOPY(FOXREF_REPLACE_MRU, THIS.aReplaceMRU)
					ENDIF
					IF TYPE("FOXREF_FOLDER_MRU") == 'C'
						=ACOPY(FOXREF_FOLDER_MRU, THIS.aFolderMRU)
					ENDIF
					IF TYPE("FOXREF_FILETYPES_MRU") == 'C'
						=ACOPY(FOXREF_FILETYPES_MRU, THIS.aFileTypesMRU)
					ENDIF

					IF TYPE("FOXREF_OPTIONS") == 'C'
						m.nCnt = ALEN(FOXREF_OPTIONS, 1)
						FOR m.i = 1 TO m.nCnt
							IF VARTYPE(FOXREF_OPTIONS[m.i, 1]) == 'C' AND PEMSTATUS(THIS, FOXREF_OPTIONS[m.i, 1], 5)
								STORE FOXREF_OPTIONS[m.i, 2] TO ("THIS." + FOXREF_OPTIONS[m.i, 1])
							ENDIF
						ENDFOR	
					ENDIF

					m.lSuccess = .T.
				ENDIF
				
				SET MEMOWIDTH TO (m.nMemoWidth)

				USE IN FoxResource
			ENDIF
		ENDIF

		SELECT (m.nSelect)
			
		RETURN m.lSuccess
	ENDFUNC


	FUNCTION UpdateLookForMRU(cPattern)
		LOCAL nRow
		
		nRow = ASCAN(THIS.aLookForMRU, cPattern, -1, -1, 1, 15)
		IF nRow > 0
			=ADEL(THIS.aLookForMRU, nRow)
		ENDIF
		=AINS(THIS.aLookForMRU, 1)
		THIS.aLookForMRU[1] = cPattern
	ENDFUNC

	FUNCTION UpdateReplaceMRU(cReplaceText)
		LOCAL nRow
		
		nRow = ASCAN(THIS.aReplaceMRU, cReplaceText, -1, -1, 1, 15)
		IF nRow > 0
			=ADEL(THIS.aReplaceMRU, nRow)
		ENDIF
		=AINS(THIS.aReplaceMRU, 1)
		THIS.aReplaceMRU[1] = cReplaceText
	ENDFUNC

	FUNCTION UpdateFolderMRU(cFolder)
		LOCAL nRow
		
		nRow = ASCAN(THIS.aFolderMRU, cFolder, -1, -1, 1, 15)
		IF nRow > 0
			=ADEL(THIS.aFolderMRU, nRow)
		ENDIF
		=AINS(THIS.aFolderMRU, 1)
		THIS.aFolderMRU[1] = cFolder
	ENDFUNC


	FUNCTION UpdateFileTypesMRU(cFileTypes)
		LOCAL nRow
		
		nRow = ASCAN(THIS.aFileTypesMRU, cFileTypes, -1, -1, 1, 15)
		IF nRow > 0
			=ADEL(THIS.aFileTypesMRU, nRow)
		ENDIF
		=AINS(THIS.aFileTypesMRU, 1)
		THIS.aFileTypesMRU[1] = cFileTypes
	ENDFUNC

	FUNCTION ProjectMatch(cFileTypes, oMatchFileCollection, cFilename)
		LOCAL i
		LOCAL lMatch

		lMatch = .F.
		FOR i = 1 TO oMatchFileCollection.Count
			IF EMPTY(JUSTPATH(oMatchFileCollection.Item(i))) OR UPPER(JUSTPATH(oMatchFileCollection.Item(i))) == UPPER(JUSTPATH(cFilename))
				IF THIS.FileMatch(JUSTFNAME(cFilename), JUSTFNAME(oMatchFileCollection.Item(i)))
					lMatch = .T.
					EXIT
				ENDIF
			ENDIF
		ENDFOR
		RETURN lMatch
	ENDFUNC

	* borrowed from Class Browser
	FUNCTION WildCardMatch(tcMatchExpList, tcExpressionSearched, tlMatchAsIs)
		LOCAL lcMatchExpList,lcExpressionSearched,llMatchAsIs,lcMatchExpList2
		LOCAL lnMatchLen,lnExpressionLen,lnMatchCount,lnCount,lnCount2,lnSpaceCount
		LOCAL lcMatchExp,lcMatchType,lnMatchType,lnAtPos,lnAtPos2
		LOCAL llMatch,llMatch2

		IF ALLTRIM(tcMatchExpList) == "*.*"
			RETURN .T.
		ENDIF

		IF EMPTY(tcExpressionSearched)
			IF EMPTY(tcMatchExpList) OR ALLTRIM(tcMatchExpList) == "*"
				RETURN .T.
			ENDIF
			RETURN .F.
		ENDIF
		lcMatchExpList=LOWER(ALLTRIM(STRTRAN(tcMatchExpList,TAB," ")))
		lcExpressionSearched=LOWER(ALLTRIM(STRTRAN(tcExpressionSearched,TAB," ")))
		lnExpressionLen=LENC(lcExpressionSearched)
		IF lcExpressionSearched==lcMatchExpList
			RETURN .T.
		ENDIF
		llMatchAsIs=tlMatchAsIs
		IF LEFTC(lcMatchExpList,1)==["] AND RIGHTC(lcMatchExpList,1)==["]
			llMatchAsIs=.T.
			lcMatchExpList=ALLTRIM(SUBSTRC(lcMatchExpList,2,LENC(lcMatchExpList)-2))
		ENDIF
		IF NOT llMatchAsIs AND " "$lcMatchExpList
			llMatch=.F.
			lnSpaceCount=OCCURS(" ",lcMatchExpList)
			lcMatchExpList2=lcMatchExpList
			lnCount=0
			DO WHILE .T.
				lnAtPos=AT_C(" ",lcMatchExpList2)
				IF lnAtPos=0
					lcMatchExp=ALLTRIM(lcMatchExpList2)
					lcMatchExpList2=""
				ELSE
					lnAtPos2=AT_C(["],lcMatchExpList2)
					IF lnAtPos2<lnAtPos
						lnAtPos2=AT_C(["],lcMatchExpList2,2)
						IF lnAtPos2>lnAtPos
							lnAtPos=lnAtPos2
						ENDIF
					ENDIF
					lcMatchExp=ALLTRIM(LEFTC(lcMatchExpList2,lnAtPos))
					lcMatchExpList2=ALLTRIM(SUBSTRC(lcMatchExpList2,lnAtPos+1))
				ENDIF
				IF EMPTY(lcMatchExp)
					EXIT
				ENDIF
				lcMatchType=LEFTC(lcMatchExp,1)
				DO CASE
					CASE lcMatchType=="+"
						lnMatchType=1
					CASE lcMatchType=="-"
						lnMatchType=-1
					OTHERWISE
						lnMatchType=0
				ENDCASE
				IF lnMatchType#0
					lcMatchExp=ALLTRIM(SUBSTRC(lcMatchExp,2))
				ENDIF
				llMatch2=THIS.WildCardMatch(lcMatchExp,lcExpressionSearched, .T.)
				IF (lnMatchType=1 AND NOT llMatch2) OR (lnMatchType=-1 AND llMatch2)
					RETURN .F.
				ENDIF
				llMatch=(llMatch OR llMatch2)
				IF lnAtPos=0
					EXIT
				ENDIF
			ENDDO
			RETURN llMatch
		ELSE
			IF LEFTC(lcMatchExpList,1)=="~"
				RETURN (DIFFERENCE(ALLTRIM(SUBSTRC(lcMatchExpList,2)),lcExpressionSearched)>=3)
			ENDIF
		ENDIF
		lnMatchCount=OCCURS(",",lcMatchExpList)+1
		IF lnMatchCount>1
			lcMatchExpList=","+ALLTRIM(lcMatchExpList)+","
		ENDIF
		FOR lnCount = 1 TO lnMatchCount
			IF lnMatchCount=1
				lcMatchExp=LOWER(ALLTRIM(lcMatchExpList))
				lnMatchLen=LENC(lcMatchExp)
			ELSE
				lnAtPos=AT_C(",",lcMatchExpList,lnCount)
				lnMatchLen=AT_C(",",lcMatchExpList,lnCount+1)-lnAtPos-1
				lcMatchExp=LOWER(ALLTRIM(SUBSTRC(lcMatchExpList,lnAtPos+1,lnMatchLen)))
			ENDIF
			FOR lnCount2 = 1 TO OCCURS("?",lcMatchExp)
				lnAtPos=AT_C("?",lcMatchExp)
				IF lnAtPos>lnExpressionLen
					IF (lnAtPos-1)=lnExpressionLen
						lcExpressionSearched=lcExpressionSearched+"?"
					ENDIF
					EXIT
				ENDIF
				lcMatchExp=STUFF(lcMatchExp,lnAtPos,1,SUBSTRC(lcExpressionSearched,lnAtPos,1))
			ENDFOR
			IF EMPTY(lcMatchExp) OR lcExpressionSearched==lcMatchExp OR ;
					lcMatchExp=="*" OR lcMatchExp=="?" OR lcMatchExp=="%%"
				RETURN .T.
			ENDIF
			IF LEFTC(lcMatchExp,1)=="*"
				RETURN (SUBSTRC(lcMatchExp,2)==RIGHTC(lcExpressionSearched,LENC(lcMatchExp)-1))
			ENDIF
			IF LEFTC(lcMatchExp,1)=="%" AND RIGHTC(lcMatchExp,1)=="%" AND ;
					SUBSTRC(lcMatchExp,2,lnMatchLen-2)$lcExpressionSearched
				RETURN .T.
			ENDIF
			lnAtPos=AT_C("*",lcMatchExp)
			IF lnAtPos>0 AND (lnAtPos-1)<=lnExpressionLen AND ;
					LEFTC(lcExpressionSearched,lnAtPos-1)==LEFTC(lcMatchExp,lnAtPos-1)
				RETURN .T.
			ENDIF
		ENDFOR
		RETURN .F.
	ENDFUNC

	* used for comparing filenames against a wildcard
	* For folder searches we can use ADIR(), but for project
	* searches we need to use this function
	FUNCTION FileMatch(cText AS String, cPattern AS String)
		LOCAL i, j, k, cPattern, nPatternLen, nTextLen, ch

		IF PCOUNT() < 2
			cPattern = THIS.cPattern
		ENDIF

		nPatternLen = LENC(cPattern)
		nTextLen = LENC(cText)

		IF nPatternLen == 0
			RETURN .T.
		ENDIF
		IF nTextLen == 0
			RETURN .F.
		ENDIF

		j = 1
		FOR i = 1 TO nPatternLen
			IF j > LENC(cText)
				RETURN .F.
			ELSE
				ch = SUBSTRC(cPattern, i, 1)
				IF ch == FILEMATCH_ANY
					j = j + 1
				ELSE
					IF ch == FILEMATCH_MORE
						FOR k = j TO nTextLen
							IF THIS.FileMatch(SUBSTRC(cText, k), SUBSTRC(cPattern, i + 1))
								RETURN .T.
							ENDIF
						ENDFOR
						RETURN .F.
					ELSE
						IF j <= nTextLen AND ch <> SUBSTRC(cText, j, 1)
							RETURN .F.
						ELSE
							j = j + 1
						ENDIF
					ENDIF
				ENDIF
			ENDIF
		ENDFOR

		* RETURN j > nTextLen
		RETURN .T.
	ENDFUNC


	* -- Parse out what's in Abstract field to return
	* -- either the search criteria or the file types searched
	FUNCTION ParseAbstract(cAbstract, cParseWhat)
		LOCAL cDisplayText
		LOCAL cSearchOptions

		cDisplayText = ''

		DO CASE
		CASE cParseWhat == "CRITERIA"
			cSearchOptions = LEFTC(cAbstract, AT_C(';', cAbstract) - 1)

			IF 'X' $ cSearchOptions
				cDisplayText = cDisplayText + IIF(EMPTY(cDisplayText), '', ", ") + COMMENTS_EXCLUDE_LOC
			ENDIF
			IF 'C' $ cSearchOptions
				cDisplayText = cDisplayText + IIF(EMPTY(cDisplayText), '', ", ") + COMMENTS_ONLY_LOC
			ENDIF
			IF 'M' $ cSearchOptions
				cDisplayText = cDisplayText + IIF(EMPTY(cDisplayText), '', ", ") + CRITERIA_MATCHCASE_LOC
			ENDIF
			IF 'W' $ cSearchOptions
				cDisplayText = cDisplayText + IIF(EMPTY(cDisplayText), '', ", ") + CRITERIA_WHOLEWORDS_LOC
			ENDIF
			IF 'P' $ cSearchOptions
				cDisplayText = cDisplayText + IIF(EMPTY(cDisplayText), '', ", ") + CRITERIA_FORMPROPERTIES_LOC
			ENDIF
			IF 'H' $ cSearchOptions AND THIS.ProjectFile <> PROJECT_GLOBAL
				cDisplayText = cDisplayText + IIF(EMPTY(cDisplayText), '', ", ") + CRITERIA_PROJECTHOMEDIR_LOC
			ENDIF
			IF 'S' $ cSearchOptions AND THIS.ProjectFile == PROJECT_GLOBAL
				cDisplayText = cDisplayText + IIF(EMPTY(cDisplayText), '', ", ") + CRITERIA_SUBFOLDERS_LOC
			ENDIF
			IF 'Z' $ cSearchOptions
				cDisplayText = cDisplayText + IIF(EMPTY(cDisplayText), '', ", ") + CRITERIA_WILDCARDS_LOC
			ENDIF
			
		CASE cParseWhat == "FILETYPES"
			cDisplayText = ALLTRIM(SUBSTRC(RTRIM(cAbstract), AT_C(';', cAbstract) + 1))
			
		ENDCASE
		
		RETURN cDisplayText
	ENDFUNC


	* -- Cleanup tables used by FoxRef -- removing
	* -- any references to files we can't find and
	* -- packing the files
	FUNCTION CleanupTables()
		LOCAL nSelect
		LOCAL cFilename
		LOCAL lSuccess
		
		m.nSelect = SELECT()
		
		m.lSuccess = .F.

		IF USED(JUSTSTEM(THIS.RefTable))
			USE IN (JUSTSTEM(THIS.RefTable))
		ENDIF
		IF USED(JUSTSTEM(THIS.DefTable))
			USE IN (JUSTSTEM(THIS.DefTable))
		ENDIF
		IF USED(JUSTSTEM(THIS.FileTable))
			USE IN (JUSTSTEM(THIS.FileTable))
		ENDIF

		IF THIS.OpenTables(.T., .T.)
			SELECT FileCursor
			DELETE ALL FOR (FileAction == FILEACTION_ERROR OR FileAction == FILEACTION_INACTIVE) AND UniqueID <> "WINDOW"
			SCAN ALL FOR !EMPTY(Filename)
				m.cFilename = ADDBS(RTRIM(FileCursor.Folder)) + RTRIM(FileCursor.Filename)
				IF !FILE(m.cFilename)
					DELETE IN FileCursor
					
					SELECT FoxDefCursor
					DELETE ALL FOR FileID == FileCursor.UniqueID IN FoxDefCursor
				ENDIF
			ENDSCAN

			SELECT FoxDefCursor
			DELETE ALL FOR Inactive IN FoxDefCursor
			
			SELECT FileCursor
			PACK IN FileCursor
			
			SELECT FoxDefCursor
			PACK IN FoxDefCursor
			
			m.lSuccess = .T.
		ELSE
			MessageBox(ERROR_EXCLUSIVE_LOC, MB_ICONEXCLAMATION, APPNAME_LOC)
		ENDIF

	
		* re-open the tables shared
		THIS.OpenTables()

		SELECT (m.nSelect)
		
		RETURN m.lSuccess
	ENDFUNC
	
	FUNCTION PrintReport(cReportName, lPreview)
		LOCAL cFilename
		LOCAL cExt
		LOCAL oReportAddIn
		
		TRY
			oReportAddIn = THIS.oReportAddIn.Item(m.cReportName)
		CATCH
			oReportAddIn = .NULL.
		ENDTRY
		
		IF !ISNULL(oReportAddIn)
			WITH oReportAddIn
				DO CASE
				CASE !EMPTY(.RptFilename)
					m.cExt = UPPER(JUSTEXT(.RptFilename))
					DO CASE
					CASE m.cExt == "PRG" OR m.cExt == "FXP"
						DO (.RptFilename) WITH m.lPreview, THIS

					CASE m.cExt == "SCX"
						DO FORM (.RptFilename) WITH m.lPreview, THIS
					
					CASE m.cExt == "FRX"
						REPORT FORM (.RptFilename)
					ENDCASE

				CASE !EMPTY(THIS.RptClassName)
				ENDCASE
			ENDWITH
		ENDIF
	ENDFUNC

	* Abstract:
	*   This program will shell out to specified file,
	*	which can be a URL (e.g. http://www.microsoft.com),
	*	a filename, etc
	*
	* Parameters:
	*	<cFile>
	*	[cParameters]
	FUNCTION ShellTo(cFile, cParameters)
		LOCAL cRun
		LOCAL cSysDir
		LOCAL nRetValue
		
		*-- GetDesktopWindow gives us a window handle to
		*-- pass to ShellExecute.
		DECLARE INTEGER GetDesktopWindow IN user32.dll
		DECLARE INTEGER GetSystemDirectory IN kernel32.dll ;
			STRING @cBuffer, ;
			INTEGER liSize

		DECLARE INTEGER ShellExecute IN shell32.dll ;
			INTEGER, ;
			STRING @cOperation, ;
			STRING @cFile, ;
			STRING @cParameters, ;
			STRING @cDirectory, ;
			INTEGER nShowCmd

		IF VARTYPE(m.cParameters) <> 'C'
			m.cParameters = ''
		ENDIF

		m.cOperation = "open"
		m.nRetValue = ShellExecute(GetDesktopWindow(), @m.cOperation, @m.cFile, @m.cParameters, '', SW_SHOWNORMAL)
		IF m.nRetValue = SE_ERR_NOASSOC && No association exists
			m.cSysDir = SPACE(260)  && MAX_PATH, the maximum path length

			*-- Get the system directory so that we know where Rundll32.exe resides.
			m.nRetValue = GetSystemDirectory(@m.cSysDir, LENC(m.cSysDir))
			m.cSysDir = SUBSTRC(m.cSysDir, 1, m.nRetValue)
			m.cRun = "RUNDLL32.EXE"
			cParameters = "shell32.dll,OpenAs_RunDLL "
			m.nRetValue = ShellExecute(GetDesktopWindow(), "open", m.cRun, m.cParameters + m.cFile, m.cSysDir, SW_SHOWNORMAL)
		ENDIF
		
		RETURN m.nRetValue
	ENDFUNC

	
ENDDEFINE

DEFINE CLASS ReportAddIn AS Custom
	Name = "ReportAddIn"
	
	ReportName      = ''
	RptFilename     = ''
	RptClassLibrary = ''
	RptClassName    = ''
	RptMethod       = ''
ENDDEFINE

DEFINE CLASS RefOption AS Custom
	Name = "RefOption"
	
	OptionName   = ''
	Description  = ''
	OptionValue  = .NULL.
	PropertyName = ''
ENDDEFINE




PROCEDURE _edgetlpos
PROCEDURE _edselect
PROCEDURE _edstopos
PROCEDURE _edgetenv
PROCEDURE _wfindtitl
PROCEDURE _wselect
