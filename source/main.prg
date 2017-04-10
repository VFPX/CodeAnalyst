LPARAMETERS tcFile,tlSilent

IF PCOUNT()=1
	tlSilent = .F.
ENDIF

IF PCOUNT()=0
	tcFile = ""
ENDIF
EXTERNAL ARRAY taArray
LOCAL llRunning
llRunning = PEMSTATUS(_SCREEN,"_Analyst",5)

IF NOT llRunning
	*- Determine if it's in fact an app running. Dont rely on the app being in the HOME()
	IF UPPER(JUSTEXT(SYS(16))) = "APP" THEN
		_SCREEN.NEWOBJECT("_Analyst","_codeAnalyzer","MAIN.PRG",SYS(16))
	ELSE
		_SCREEN.NEWOBJECT("_analyst","_codeAnalyzer","MAIN.PRG")
	ENDIF
	llRunning = .T.
ENDIF
_SCREEN._Analyst.IsDirectory = .F.
_SCREEN._Analyst.cdirectory = ""
IF UPPER(tcFile)="DIRECTORY"
	IF DIRECTORY(tlSilent)
	_SCREEN._Analyst.IsDirectory = .T.
	_SCREEN._Analyst.cDirectory = tlSilent
	 _SCREEN._Analyst.BuildFakeProject(tlSilent)
	 tcFile = _SCREEN._Analyst.cFile
	 tlSilent = .T.
	 ELSE
	 MESSAGEBOX("Directory " + tlSilent + " does not exist",0+16,"Code Analyst")
	 tlSilent = .F.
	 tcFile = ""
	 
	ENDIF
ENDIF
IF tlSilent
	_SCREEN._Analyst.lProjectRun = .T.
ENDIF
*- Check for a parameter passed. If none is passed, the analyst is just installed in the menu
*- If the analyst was already installed, start the Analyze anyway.


IF PCOUNT()>0 AND UPPER(tcFile)="-CONFIG"
	_SCREEN._Analyst.Configure()
ELSE
	IF PCOUNT() = 1 OR llRunning THEN
		_SCREEN._Analyst.Analyze(tcFile)
	ENDIF
ENDIF

DEFINE CLASS _codeAnalyzer AS CUSTOM
	cFile = ""
	csetesc = ""
	csetesclabel = ""
	cMainProgram = ""
	cHomeDir = ""
	isdirectory = .F.
	cDirectory = ""
	lLineRules = .F.
	cLine = ""
	nLine = 0
	cFuncName = ""
	oTherm = .NULL.
	oObject = .NULL.
	cObject= ""
	cCode = ""
	cFontString = "Tahoma,8,N"
	cError = ""
	cHomeDir = ""
	cRuleDir = ""
	cResetFile = ""
	cAnalysisCursor = ""
	nFuncLines = 0
	cWarningID = ""
	nFileLines = 0
	cMessage = ""
	lDisplayMessage = .T.
	lDisplayForm = .T.
	lUseDefaultDir = .T.
	lProjectRun = .F.
	cTable = ""
	cClassName = ""


	PROCEDURE RESET
		IF NOT EMPTY(THIS.cAnalysisCursor)
			IF USED(THIS.cAnalysisCursor)
				LOCAL lc
				lc = SET("SAFETY")
				SET SAFETY OFF
				ZAP IN (THIS.cAnalysisCursor)
				SET SAFETY &lc
			ENDIF
		ENDIF
	ENDPROC

	PROCEDURE SetPrefs
		LOCAL nSelect
		LOCAL lcRes
		lcRes = "ANALYST"
		LOCAL lSuccess
		LOCAL nMemoWidth
		LOCAL nCnt
		LOCAL cData

		_AnalystResetXML = THIS.cResetFile
		_AnalystFontString       = THIS.cFontString
		_Analysthomedir          = THIS.cHomeDir
		_AnalystRuledir          = THIS.cRuleDir
		IF THIS.lUseDefaultDir
			_AnalystRuledir = ""
		ENDIF

		nSelect = SELECT()

		lSuccess = .F.

		* make sure Resource file exists and is not read-only
		nCnt = ADIR(aFileList, SYS(2005))
		IF nCnt > 0 AND ATC('R', aFileList[1, 5]) == 0
			USE (SYS(2005)) IN SELECT("FOXRESOURCE") SHARED AGAIN ALIAS FoxResource
			IF USED("FoxResource") AND !ISREADONLY("FoxResource")
				nMemoWidth = SET('MEMOWIDTH')
				SET MEMOWIDTH TO 255

				SELECT FoxResource
				LOCATE FOR UPPER(ALLTRIM(TYPE)) == "PREFW" AND UPPER(ALLTRIM(ID)) == lcRes AND EMPTY(NAME)
				IF !FOUND()
					APPEND BLANK IN FoxResource
					REPLACE ;
						TYPE WITH "PREFW", ;
						ID WITH lcRes, ;
						READONLY WITH .F. ;
						IN FoxResource
				ENDIF

				IF !FoxResource.READONLY

					SAVE TO MEMO DATA ALL LIKE _Analyst*

					REPLACE ;
						UPDATED WITH DATE(), ;
						ckval WITH VAL(SYS(2007, FoxResource.DATA)) ;
						IN FoxResource

					lSuccess = .T.
				ENDIF
				SET MEMOWIDTH TO (nMemoWidth)

				USE IN FoxResource
			ENDIF
		ENDIF

		SELECT (nSelect)

		RETURN lSuccess

	ENDPROC

	PROCEDURE GetPrefs
		LOCAL nSelect
		LOCAL lcRes
		lcRes = "ANALYST"
		LOCAL lSuccess
		LOCAL nMemoWidth


		nSelect = SELECT()

		lSuccess = .F.
		IF EMPTY(This.cRuleDir)
			This.cRuleDir = HOME()
		ENDIF

		IF FILE(SYS(2005))    && resource file not found.
			USE (SYS(2005)) IN SELECT("FOXRESOURCE") SHARED AGAIN ALIAS FoxResource
			IF USED("FoxResource")
				nMemoWidth = SET('MEMOWIDTH')
				SET MEMOWIDTH TO 255

				SELECT FoxResource
				LOCATE FOR UPPER(ALLTRIM(TYPE)) == "PREFW" ;
					AND UPPER(ALLTRIM(ID)) == lcRes ;
					AND !DELETED()

				IF FOUND() AND !EMPTY(DATA) AND ckval == VAL(SYS(2007, DATA)) AND EMPTY(NAME)
					RESTORE FROM MEMO DATA ADDITIVE

					THIS.cFontString = _AnalystFontString
					IF TYPE("_ANALYSTRuleDIR")="C"
						IF NOT EMPTY(_AnalystRuledir)
							THIS.lUseDefaultDir = .F.
							THIS.cRuleDir = _AnalystRuledir
						ELSE
							THIS.lUseDefaultDir = .T.
							THIS.cRuleDir = CURDIR()
						ENDIF
					ELSE
						THIS.lUseDefaultDir = .T.
						THIS.cRuleDir = CURDIR()
					ENDIF
					IF TYPE("_ANALYSTHomeDIR")="C"
						IF NOT EMPTY(_Analysthomedir)
						ELSE
							THIS.cHomeDir = CURDIR()
						ENDIF
					ELSE
						THIS.cHomeDir = CURDIR()
					ENDIF



					IF TYPE("_AnalystResetXML")="C"
						IF NOT EMPTY(_AnalystResetXML)
							THIS.cResetFile = _AnalystResetXML
						ELSE
							THIS.cResetFile = ""
						ENDIF
					ELSE
						THIS.cResetFile = ""
					ENDIF


					lSuccess = .T.
				ENDIF

				SET MEMOWIDTH TO (nMemoWidth)

				USE IN FoxResource
			ENDIF
		
		ENDIF

		SELECT (nSelect)

		RETURN lSuccess

	ENDPROC
	PROCEDURE CreateRuleTable

		IF NOT FILE(THIS.cRuleDir+"CODERULE.DBF")
			LOCAL lnArea
			lnArea = SELECT()
			SELECT 0
			CREATE TABLE (THIS.cRuleDir+"CODERULE.DBF") (;
				TYPE C(1),;
				NAME C(30),;
				ACTIVE L,;
				DESCRIPT M,;
				Script M,;
				PROGRAM M,;
				CLASSLIB C(30),;
				Classname C(50),;
				TIMESTAMP T,;
				UniqueID C(10);
				)

			IF NOT USED("_CODERULE")
				USE _CODERULE IN 0
			ENDIF

			SELECT _CODERULE
			SCAN
				SCATTER MEMVAR MEMO
				INSERT INTO (THIS.cRuleDir+"CODERULE") FROM MEMVAR
			ENDSCAN
			USE IN SELECT("_CODERULE")

		ENDIF

	ENDPROC

	PROCEDURE DESTROY
		IF NOT ISNULL(THIS.oTherm)
			THIS.oTherm.HIDE()
			THIS.oTherm.RELEASE()
		ENDIF
		IF NOT EMPTY(PRMBAR("_MTOOLS", 5942)) THEN
			RELEASE BAR 5942 OF _MTOOLS
		ENDIF
	ENDPROC

	PROCEDURE Configure
		LOCAL lnArea
		lnArea = SELECT()
		IF NOT USED("CODERULE")
			*- Use the cHomeDir property here... again don't rely on the HOME()
			USE (THIS.cRuleDir+"CODERULE") IN 0
		ENDIF
		DO FORM ConfigureAnalyst
		SELECT (lnArea)
	ENDPROC
	PROCEDURE AddWarning
		LPARAMETERS tcWarning,tcType
		IF NOT EMPTY(THIS.aWarnings(1,1))
			DIMENSION THIS.aWarnings(ALEN(THIS.aWarnings,1)+1,3)
		ENDIF
		DIMENSION THIS.aWarnings(ALEN(THIS.aWarnings,1),3)
		THIS.aWarnings(ALEN(THIS.aWarnings,1),1)=tcWarning
		IF EMPTY(THIS.cFile)
			THIS.cFile = "Unknown"
		ENDIF
		THIS.aWarnings(ALEN(THIS.aWarnings,1),2)=THIS.cFile
		THIS.aWarnings(ALEN(THIS.aWarnings,1),3)=THIS.cWarningID
		IF USED(THIS.cAnalysisCursor)
			LOCAL lnArea
			lnArea = SELECT()
			SELECT (THIS.cAnalysisCursor)
			THIS.AddWarningCursor(tcWarning)
			*!* REPLACE warnings WITH warnings + tcWarning +CHR(13)+CHR(10)
			SELECT (lnArea)
		ENDIF
	ENDPROC

	PROCEDURE AddWarningCursor
		LPARAMETERS tcWarning

		LOCAL lcFile,lcID,lcFunc,lcType
		lcType = "Warning"
		lcID = THIS.cWarningID
		lcFile = THIS.cFile
		lcFunc = THIS.cFuncName
		lnLine = THIS.nLine
		IF NOT EMPTY(THIS.cObject)
			lcFunc = TRIM(THIS.cObject) + "."+lcFunc
		ENDIF
		IF EMPTY(THIS.cAnalysisCursor)
			THIS.BuildAnalysisCursor()
		ENDIF
		lcFileType = THIS.GetFileType()

		INSERT INTO (THIS.cAnalysisCursor) (cfunc,cprog,cType,cFileType,cWarning,nLine,warnings) ;
			VALUES (lcFunc,lcFile,lcType,lcFileType,lcID,lnLine,tcWarning)

	ENDPROC

	PROCEDURE GetFileType
		LOCAL lcExt,lcRet
		lcRet = "Code"
		lcExt = UPPER(JUSTEXT(THIS.cFile))
		DO CASE
			CASE lcExt = "VCX"
				lcRet = "Classes"
			CASE lcExt = "SCX"
				lcRet = "Forms"
			CASE lcExt = "MNX"
				lcRet = "Menus"
			CASE lcExt = "PRG"
				lcRet = "Programs"
			CASE lcExt = "APP"
				lcRet = "Apps"

		ENDCASE
		RETURN lcRet
	ENDPROC

	PROCEDURE AddMessage
		LPARAMETERS tcMsg
		THIS.cMessage = THIS.cMessage + IIF(EMPTY(THIS.cMessage),"",CHR(13)+CHR(10))+tcMsg
	ENDPROC

	PROCEDURE INIT
		*- Again don't rely on the HOME(), use SYS(16) instead. Now strip class/method names to
		*- get the homedir for the data
		LOCAL lcProgram
		LOCAL lcSetExact
		lcSetExact=SET("EXACT")
		SET EXACT OFF
		IF SYS(16)="PROCEDURE"
			lcProgram = ALLTRIM(SUBSTR(SYS(16),ATC(" ",SYS(16),2)+1))
		ELSE
			lcProgram = ALLTRIM(STREXTRACT(SYS(16), " ", " ", 2, 2))
		ENDIF
		THIS.cHomeDir = JUSTPATH(lcProgram)+"\"

		IF NOT PEMSTATUS(THIS,"aCode",5)
			THIS.ADDPROPERTY("aCode(1,4)")
		ENDIF
		IF NOT PEMSTATUS(THIS,"awarnings",5)
			THIS.ADDPROPERTY("awarnings(1,3)")
		ENDIF
		THIS.aWarnings(1,1) = ""
		THIS.aWarnings(1,2) = ""
		THIS.aWarnings(1,3) = ""
		THIS.aCode(1,1) = ""
		IF NOT PEMSTATUS(THIS,"aRules",5)
			THIS.ADDPROPERTY("aRules(1,5)")
		ENDIF
		THIS.GetPrefs()
		THIS.CreateRuleTable()

		THIS.LoadRules()
		LOCAL lcDir
		lcDir = "DO ('"+THIS.cHomeDir +"ANALYST.APP')"
		
		
		DEFINE BAR 5942 OF _MTOOLS PROMPT "Code Analyst..." AFTER  _MTL_TOOLBOX
		ON SELECTION BAR 5942 OF _MTOOLS &lcDir

		SET EXACT &lcSetExact

	ENDPROC

	PROCEDURE BuildAnalysisCursor
		LOCAL lnArea
		lnArea = SELECT()
		SELECT 0
		THIS.cAnalysisCursor = SYS(2015)
		CREATE CURSOR (THIS.cAnalysisCursor) (;
			cFileType C(10),;
			cfunc C(50),;
			cprog C(125),;
			cClass C(50),;
			cType C(10),;
			nLine N(6),;
			cWarning C(10),;
			warnings M;
			)

		LOCAL lc
		lc = SET("COLLATE", "TO")
		SET COLLATE TO "MACHINE"

		INDEX ON cfunc+cprog+cClass TAG funcProg

		SET COLLATE TO (lc)

		SELECT (lnArea)

	ENDPROC

	PROCEDURE AddToCursor
		LPARAMETERS tcFunc, tcProg, tcClass,tcType
		*% Add tcClass
		IF EMPTY(tcType)
			tcType = JUSTEXT(tcProg)
		ENDIF
		IF EMPTY(tcType)
			tcType = "Unknown"
		ENDIF
		IF EMPTY(THIS.cAnalysisCursor)
			THIS.BuildAnalysisCursor()
		ELSE
			IF NOT USED(THIS.cAnalysisCursor)
				THIS.BuildAnalysisCursor()
			ENDIF
		ENDIF
		IF EMPTY(tcFunc)
			tcFunc = THIS.cFile
		ENDIF
		IF EMPTY(tcProg)
			tcProg = THIS.cFuncName
		ENDIF
		IF EMPTY(tcClass)
			tcClass = THIS.cClassName
		ENDIF
		tcFunc = PADR(tcFunc,50)
		tcProg = PADR(tcProg,125)
		IF NOT SEEK(tcFunc+tcProg+tcClass,THIS.cAnalysisCursor)

			INSERT INTO (THIS.cAnalysisCursor) (cfunc,cprog,cClass,cType) ;
				VALUES (tcFunc,tcProg, tcClass, tcType)
		ENDIF
	ENDPROC

	PROCEDURE LoadRules

		IF FILE(THIS.cRuleDir+"CODERULE.DBF")
			SELECT NAME,TYPE,Script,UniqueID FROM THIS.cRuleDir+"CODERULE" WHERE ACTIVE INTO ARRAY THIS.aRules
			LOCAL lni
			FOR lni = 1 TO ALEN(THIS.aRules,1)
				IF EMPTY(THIS.aRules(lni,1))
					LOOP
				ENDIF
				IF THIS.aRules(lni,2)="L"
					THIS.lLineRules = .T. && Needed for performance checks
					EXIT
				ENDIF
			ENDFOR
		ENDIF

	ENDPROC

	PROCEDURE ResetArrays
		DIMENSION THIS.aCode(1,4)
		THIS.aCode(1,1) = ""
		DIMENSION THIS.aWarnings(1,3)
		THIS.aWarnings(1,1) = ""
		THIS.aWarnings(1,2) = ""
		THIS.aWarnings(1,3) = ""

	ENDPROC

	PROCEDURE Analyze
		LPARAMETERS tcFile
		THIS.cMessage = ""
		THIS.cError = ""
		LOCAL lcDir
		lcDir = CURDIR()
		THIS.cMainProgram = tcFile
		IF NOT EMPTY(tcFile)
			IF NOT FILE(tcFile) AND NOT USED(tcFile)
				MESSAGEBOX("File "+tcFile+" does not exist.",16,"Code Analyst")
				RETURN

			ENDIF
		ENDIF
		IF ISNULL(THIS.oTherm)
			THIS.oTherm = NEWOBJECT("cprogressform","foxref.vcx",THIS.cHomeDir+"ANALYST.APP")
			THIS.oTherm.SetMax(100)
		ENDIF
		LOCAL lc
		THIS.ResetArrays()
		LOCAL lAlias 
		lAlias = .F.
		IF EMPTY(tcFile)
			IF ASELOBJ(la,1)=0
				** Should we use the Active project or not?
				*!* IF TYPE("_VFP.ActiveProject")="U"

				tcFile = GETFILE("PRG;PJX;VCX;SCX","Select file","Open",0,"Select file to analyze")

				*!*					ELSE
				*!*						tcFile = _VFP.ActiveProject.Name
				*!*					ENDIF
				IF EMPTY(tcFile)
					RETURN
				ENDIF
				SET DEFAULT TO (STRTRAN(tcFile,JUSTFNAME(tcFile)))
				THIS.cHomeDir = (STRTRAN(tcFile,JUSTFNAME(tcFile)))


			ENDIF
		ELSE
			IF USED(tcFile)
				lAlias = .T.
			ELSE
				SET DEFAULT TO (STRTRAN(tcFile,JUSTFNAME(tcFile)))
				THIS.cHomeDir = (STRTRAN(tcFile,JUSTFNAME(tcFile)))
			ENDIF
		ENDIF

		TRY
			THIS.csetesc = ON("ESCAPE")
			THIS.csetEscLabel = ON("KEY","ESC")
			ON ESCAPE _SCREEN._Analyst.StopAnalysis()
			ON KEY LABEL ESCAPE _SCREEN.StopAnalysis = .T.

			THIS.PreValidate()
			IF EMPTY(tcFile)
				lc = "Current Object"
				THIS.cFile = lc
				THIS.cMainProgram = lc
				THIS.oTherm.SetDescription("Analyzing "+lc)
				THIS.oTherm.SetProgress(1)
				DOEVENTS
				THIS.oTherm.SHOW()
				THIS.AddToCursor(la(1).NAME,la(1).NAME,'',"Object")

				THIS.AnalyzeCurrObj()
			ELSE
				lc = tcFile
				IF EMPTY(THIS.cMainProgram)
					THIS.cMainProgram = lc
				ENDIF
				THIS.cFile = lc
				THIS.oTherm.SetDescription("Analyzing "+lc)
				THIS.oTherm.SetProgress(1)
				DOEVENTS
				THIS.oTherm.SHOW()
				THIS.AddToCursor(lc,lc,'',"File")
				THIS.AnalFile(tcFile,lAlias)
			ENDIF

			THIS.oTherm.HIDE()
			IF THIS.lDisplayForm
				IF NOT THIS.lProjectRun OR RECCOUNT(THIS.cAnalysisCursor)>0
					IF NOT PEMSTATUS(_SCREEN,"_analysisform",5)
						_SCREEN.ADDPROPERTY("_analysisform",.NULL.)
					ENDIF
					IF ISNULL(_SCREEN._Analysisform)
						SELECT 0
						DO FORM codeanalresults NAME _SCREEN._Analysisform
					ELSE
						_SCREEN._Analysisform.REFRESH()
					ENDIF
				ENDIF
			ENDIF
		CATCH TO loErr
			IF NOT ISNULL(THIS.oTherm)
				THIS.oTherm.HIDE()
			ENDIF

			ERROR loErr.MESSAGE+ " on line " + TRANSFORM(loErr.LINENO)+" of "+loErr.PROCEDURE
		ENDTRY
		THIS.PostValidate()
		SET DEFAULT TO (lcDir)
		LOCAL lc
		lc = THIS.csetesc
		ON ESCAPE &lc
		lc = THIS.cSetEscLabel
		ON KEY LABEL ESCAPE &lc


	PROCEDURE StopAnalysis
		LOCAL lc
		IF PEMSTATUS(_SCREEN,"StopAnalysis",5)
			_SCREEN.StopAnalysis = .T.
		ENDIF
		lc = THIS.csetesc
		ON ESCAPE &lc
		THIS.oTherm.HIDE()
		CANCEL

	PROCEDURE AnalyzeCurrObj
		LPARAMETERS tlWork

		THIS.cFile = "Current Object: "+la(1).NAME
		DIMENSION THIS.aCode(1,4)
		THIS.aCode(1,1) = ""

		THIS.analyzeObj(la(1))
		lc = "Code Review" + CHR(13)+CHR(10)

		IF NOT EMPTY(THIS.aCode(1,1))
			=ASORT(THIS.aCode,2)

		ENDIF
		lc = lc + CHR(13)+CHR(10)+"--- Good Code ---"+CHR(13)+CHR(10)
		LOCAL llTitle
		llTitle = .F.
		FOR lni = 1 TO ALEN(THIS.aCode,1)
			IF EMPTY(THIS.aCode(lni,1))
				LOOP
			ENDIF
			IF THIS.aCode(lni,2)>40 AND NOT llTitle
				llTitle = .T.
				lc = lc + CHR(13)+CHR(10)+"--- Possible Candidates ---"+CHR(13)+CHR(10)
			ENDIF
			IF NOT tlWork  OR (tlWork AND ;
					(THIS.aCode(lni,2)>40 OR ;
					(THIS.aCode(lni,3)<(MAX(1,THIS.aCode(lni,2)/2)) AND NOT ;
					THIS.aCode(lni,3)=THIS.aCode(lni,2))))

			ELSE
				THIS.aCode(lni,1) = "DELETE"
			ENDIF
		ENDFOR

		ln = ALEN(THIS.aCode,1)
		ln2 = 1
		FOR lni = 1 TO ln
			IF NOT EMPTY(THIS.aCode(lni,1))
				IF THIS.aCode(lni,1)="DELETE"
					=ADEL(THIS.aCode,lni)
					lni = lni-1
					LOOP
				ELSE
					ln2=ln2+1
				ENDIF
			ENDIF
		ENDFOR
		DIMENSION THIS.aCode(ln2,4)
		IF EMPTY(THIS.aCode(ln2,1))
			THIS.aCode(ln2,1)=""
		ENDIF
		RETURN


	PROCEDURE analyzeObj
		LPARAMETERS toObj

		LOCAL lni
		LOCAL lcText
		lcText = ""
		LOCAL loObj

		LOCAL la(1)
		LOCAL lnMethods
		THIS.oTherm.setstatus("Object: "+toObj.NAME)

		lnMethods = AMEMBERS(la,toObj,1)

		THIS.AddToCursor(toObj.NAME,toObj.NAME,'',"Object")

		IF NOT PEMSTATUS(toObj,"Name",5)
			RETURN
		ENDIF

		FOR lni = 1 TO lnMethods
			IF la(lni,2)="M" OR la(lni,2)="E"
				IF PEMSTATUS(toObj,"ReadMethod",5)
					lcContent = toObj.READMETHOD(la(lni,1))
					IF NOT EMPTY(lcContent)
						THIS.Add2Array(toObj.NAME+"."+la(lni,1),ALINES(laX,lcContent),@laX)
						*!* lcText = lcText + toobj.Name+"."+la(lni,1) + " - "+LTRIM(STR(ALINES(laX,lcCOntent)))+CHR(13)+CHR(10)
					ENDIF
				ENDIF
			ENDIF
		ENDFOR
		THIS.ValidateObject(toObj)

		FOR lni = 1 TO lnMethods
			IF la(lni,2)="Object"
				loObj = toObj.&la(lni,1)
				THIS.analyzeObj(loObj)
			ENDIF
		ENDFOR


	PROCEDURE AnalyzeCode
		LPARAMETERS tcFile,tlWork
		IF EMPTY(tcFile)
			tcFile = GETFILE("PRG")
		ENDIF

		DIMENSION THIS.aCode(1,4)
		THIS.aCode(1,1) = ""
		THIS.aCode(1,2) = 0
		THIS.aCode(1,3) = 0
		THIS.aCode(1,4) = ""
		LOCAL lni
		IF MEMLINES(tcFile)>1
			THIS.analstring(tcFile)

		ELSE
			THIS.AnalFile(tcFile)
		ENDIF

		lc = "Code Review" + CHR(13)+CHR(10)

		=ASORT(THIS.aCode,2)

	ENDPROC

	PROCEDURE ScanSCXVCX
		LPARAMETERS tcFile
		THIS.cFile = tcFile
		LOCAL lnArea
		lnArea = SELECT()
		LOCAL lcAlias
		lcAlias = SYS(2015)
		SELECT 0
		USE (tcFile) AGAIN SHARED ALIAS &lcAlias
		SCAN FOR NOT EMPTY(methods)
			THIS.cObject = TRIM(PARENT)+IIF(EMPTY(PARENT),"",".")+TRIM(objname)
			THIS.cClassName = objname
			THIS.analstring(methods,tcFile)
			THIS.cObject = ""
			THIS.cClassName = ""
		ENDSCAN
		SELECT (lcAlias)
		USE
		SELECT (lnArea)
	ENDPROC

	PROCEDURE ScanMNX
		LPARAMETERS tcFile
		THIS.cFile = tcFile
		LOCAL lnArea
		lnArea = SELECT()
		LOCAL lnArea,lcAlias
		lcAlias = SYS(2015)
		SELECT 0
		USE (tcFile) AGAIN SHARED ALIAS &lcAlias
		SCAN
			IF NOT EMPTY(SETUP)
				THIS.analstring(SETUP,tcFile+" Setup")
			ENDIF
			IF NOT EMPTY(PROCEDURE)
				THIS.analstring(PROCEDURE,tcFile + " Procedures")
			ENDIF
			IF NOT EMPTY(cleanup)
				THIS.analstring(cleanup,tcFile + " Cleanup")
			ENDIF
		ENDSCAN
		SELECT (lcAlias)
		USE
		SELECT (lnArea)
	ENDPROC

	PROCEDURE BuildFakeProject
		LPARAMETERS tcDir
		LOCAL lnArea
		lnArea = SELECT()
		LOCAL lcFile
		lcFile = SYS(2015)
		THIS.cFile = lcFile
		SELECT 0
		CREATE CURSOR (lcFile) (TYPE C(1), NAME M)
		THIS.addtoproj(tcDir,lcFile)
		SELECT (lnArea)

	ENDPROC
	FUNCTION GetFileExtensionType
		LPARAMETERS tcExt
		LOCAL lcRet
		lcRet = "X"
		DO CASE
			CASE tcExt="PRG"
				lcRet = "P"
		ENDCASE

		RETURN lcRet
	ENDFUNC

	PROCEDURE AddToProj
		LPARAMETERS tcDir,tcAlias
		LOCAL lnFiles
		LOCAL lnDirs
		LOCAL la(1)
		LOCAL lni
		LOCAL lcExt

IF RIGHT(tcDir,1)="\"
	tcDir = LEFT(tcDir,LEN(tcDir)-1)
ENDIF
		lnFiles = ADIR(la,tcDir+"\*.*")
		FOR lni = 1 TO lnFiles
			IF LEFT(la(lni,1),1)<>"."
				lcExt = JUSTEXT(la(lni,1))
				IF INLIST(lcExt,"PRG","VCX","MNX","FRX","SCX")
					INSERT INTO (tcAlias) VALUES ("P",FULLPATH(tcDir+"\"+la(lni,1)))
				ENDIF
			ENDIF
		ENDFOR

		lnDirs = ADIR(la,tcDir+"\*.","D")
		FOR lni = 1 TO lnDirs
			IF LEFT(la(lni,1),1)<>"."
				
				THIS.AddToProj(tcDir+"\"+la(lni,1), tcAlias)
			ENDIF
		ENDFOR
	ENDPROC
	
	PROCEDURE scanDirectory
		LPARAMETERS tcFile
		
		LOCAL lnArea,lcFile
		lnArea = SELECT()
		LOCAL lnArea,lcAlias
		IF NOT PEMSTATUS(_SCREEN,"StopAnalysis",5)
			_SCREEN.AddProperty("StopAnalysis",.F.)
		ENDIF
		_SCREEN.StopAnalysis = .F.
		SELECT (tcFile)
		lcAlias = ALIAS()
		SCAN FOR NOT TYPE="H"
			IF _Screen.StopAnalysis
				EXIT
			ENDIF
			THIS.oTherm.SetProgress(RECNO()/RECCOUNT()*95)
			lcFile = STRTRAN(NAME,CHR(0))
			THIS.AnalFile(lcFile)
		ENDSCAN
		SELECT (lcAlias)
		USE
		SELECT (lnArea)

	ENDPROC

	PROCEDURE scanProject
		LPARAMETERS tcFile
		THIS.cFile = tcFile
		LOCAL lnArea,lcFile
		lnArea = SELECT()
		LOCAL lnArea,lcAlias
		lcAlias = SYS(2015)
		SELECT 0
		USE (tcFile) AGAIN SHARED ALIAS &lcAlias
		SCAN FOR NOT TYPE="H"
			THIS.oTherm.SetProgress(RECNO()/RECCOUNT()*95)
			lcFile = STRTRAN(NAME,CHR(0))
			THIS.AnalFile(lcFile)
		ENDSCAN
		SELECT (lcAlias)
		USE
		SELECT (lnArea)

	ENDPROC
	PROCEDURE AnalFile
		LPARAMETERS tcFile,tlAlias

IF PCOUNT()=1
	tlAlias = .F.
ENDIF
		THIS.oTherm.SetDescription("Analyzing "+tcFile)
		THIS.cFile = tcFile
IF tlAlias
THIS.ScanDirectory(tcFile)
ELSE
LOCAL lcRet
lcRet = ""

		LOCAL lcExt
		lcExt = UPPER(JUSTEXT(tcFile))
		TRY
		DO CASE
			CASE lcExt = "PRG"
				lcRet = THIS.analstring(FILETOSTR(tcFile))
			CASE lcExt = "SCX"
				lcRet = THIS.ScanSCXVCX(tcFile)
			CASE lcExt = "MNX"
				lcRet = THIS.ScanMNX(tcFile)

			CASE lcExt = "FRX"
				lcRet = ""

			CASE lcExt="ZIP"
				lcRet = "Zip File - Ignored"

			CASE lcExt = "BAK"
				lcRet = "Ignored"
			CASE lcExt = "APP"
				lcRet = ""

			CASE lcExt = "FLL"
				lcRet = ""

			CASE lcExt = "DBF"
				lcRet = ""

			CASE lcExt = "PJX"
				lcRet = THIS.scanProject(tcFile)

			CASE lcExt = "VCX"
				lcRet = THIS.ScanSCXVCX(tcFile)

			CASE INLIST(lcExt,"BMP","TXT","MSK","INC","H","JPG","GIF","ICO","SCT","MNT","FPT","TBK","PJT","VCT","FRA","SCA","MNA","VCA","XML","HTM","FXP")

			OTHERWISE
			LOCAL lc
				TRY
				lc = THIS.analstring(FILETOSTR(tcFile))
				CATCH
				lc = ""
				ENDTRY
				lcRet = lc
		ENDCASE
		
		CATCH TO loErr
		lcRet = "Error reading: " + tcFile +": "+loErr.Message
		ENDTRY
		RETURN lcRet
ENDIF

	PROCEDURE analstring
		** Takes a piece of code and looks for any breaks in it.
		** if someone wanted to write a rule to analyze entire pieces of code, here is where
		** it would go.
		** so the first rule is to break it into individual lines and analyze them.

		LPARAMETERS tcString,tcName

		LOCAL la(1)
		LOCAL laBreak(3)
		laBreak(1) = "PROCEDURE"
		laBreak(2) = "FUNCTION"
		laBreak(3) = "DEFINE"

		LOCAL lnTotal
		lnTotal=ALINES(la,tcString)

		LOCAL lni

		LOCAL lcText
		lcText = ""
		LOCAL lcFunc,lcWord
		LOCAL lnCount
		lnCount = 0
		LOCAL laX(1)
		laX(1)= ""
		lcFunc = "Program" && JUSTFNAME(tcFile)
		FOR lni = 1 TO lnTotal
			THIS.oTherm.setstatus("Line "+LTRIM(STR(lni))+" of "+LTRIM(STR(lnTotal)))
			lcText = la(lni)
			*% Don't uppercase - some rules may need to be case sensitive
			*% lcText = ALLTRIM(UPPER(STRTRAN(lcText,"	")))
			lcText = ALLTRIM(STRTRAN(lcText,"	"))
			IF EMPTY(lcText)
				LOOP
			ENDIF
			*!* THIS.ValidateLine(lcText)
			lcWord = LEFT(lcText,ATC(" ",lcText)-1)
			IF EMPTY(lcWord) OR ASCAN(laBreak,UPPER(lcWord))=0
				lnCount = lnCount+1
				DIMENSION laX(IIF(EMPTY(laX(1)),1,ALEN(laX,1)+1))
				laX(ALEN(laX,1))=lcText
			ELSE
				THIS.Add2Array(lcFunc,lnCount,@laX)
				lnCount = 0
				IF lcText = "DEFINE CLASS"
					lcFunc = ALLTRIM(STRTRAN(lcText,"DEFINE CLASS"))
					lcFunc = LEFT(lcFunc,ATC(" ",lcFunc)-1)
					THIS.cObject = lcFunc
				ELSE
					lcFunc = ALLTRIM(STRTRAN(lcText,lcWord))
				ENDIF
				DIMENSION laX(1)
				laX(1) = ""
			ENDIF
		ENDFOR
		THIS.Add2Array(lcFunc,lnCount,@laX)
		THIS.nFileLines = lnTotal
		THIS.AddToCursor(tcName,tcName,THIS.cClassName,"File")
		THIS.ValidateFile(tcString,tcName)

	ENDPROC

	PROCEDURE ValidateObject
		LPARAMETERS toObj
		LOCAL lni,lcFunc
		THIS.oObject = toObj
		THIS.cObject = toObj.NAME
		FOR lni = 1 TO ALEN(THIS.aRules,1)
			IF EMPTY(THIS.aRules(lni,1))
				LOOP
			ENDIF
			IF ALEN(THIS.aRules,2)>3
				THIS.cWarningID = THIS.aRules(lni,4)
			ELSE
				THIS.cWarningID = ""

			ENDIF
			IF THIS.aRules(lni,2)="O"
				lcFunc = THIS.aRules(lni,3)
				EXECSCRIPT(lcFunc)
			ENDIF
		ENDFOR

	ENDPROC
	PROCEDURE ValidateLine
		LPARAMETERS tcLine
		LOCAL lni,lcFunc

		THIS.cLine = tcLine
		FOR lni = 1 TO ALEN(THIS.aRules,1)
			IF EMPTY(THIS.aRules(lni,1))
				LOOP
			ENDIF
			IF ALEN(THIS.aRules,2)>3
				THIS.cWarningID = THIS.aRules(lni,4)
			ELSE
				THIS.cWarningID = ""

			ENDIF
			IF THIS.aRules(lni,2)="L"
				lcFunc = THIS.aRules(lni,3)
				EXECSCRIPT(lcFunc)
			ENDIF
		ENDFOR

	PROCEDURE ValidateFile
		LPARAMETERS tcFile,tcName

		LOCAL lni,lcFunc
		THIS.cFile = tcName
		THIS.cCode = tcFile
		FOR lni = 1 TO ALEN(THIS.aRules,1)
			IF EMPTY(THIS.aRules(lni,1))
				LOOP
			ENDIF
			IF ALEN(THIS.aRules,2)>3
				THIS.cWarningID = THIS.aRules(lni,4)
			ELSE
				THIS.cWarningID = ""

			ENDIF
			IF THIS.aRules(lni,2)="F"
				lcFunc = THIS.aRules(lni,3)
				EXECSCRIPT(lcFunc)
			ENDIF
		ENDFOR

	PROCEDURE ValidateCode
		LPARAMETERS tcCode,tcName

		THIS.cFuncName = tcName
		IF EMPTY(tcName) OR tcName="Program"
			THIS.cFuncName = THIS.cFile
		ENDIF
		THIS.oTherm.setstatus("Function: "+tcName)
		LOCAL lni,lcFunc
		THIS.cCode = tcCode
		FOR lni = 1 TO ALEN(THIS.aRules,1)
			IF EMPTY(THIS.aRules(lni,1))
				LOOP
			ENDIF
			IF ALEN(THIS.aRules,2)>3
				THIS.cWarningID = THIS.aRules(lni,4)
			ELSE
				THIS.cWarningID = ""

			ENDIF
			IF THIS.aRules(lni,2)="M"
				lcRule = THIS.aRules(lni,1)
				lcFunc = THIS.aRules(lni,3)
				TRY
					EXECSCRIPT(lcFunc)
				CATCH TO loErr
					THIS.AddError(loErr,lcRule,lcFunc)
					THIS.aRules(lni,1) = ""
				ENDTRY
			ENDIF
		ENDFOR


	PROCEDURE PreValidate

		THIS.oTherm.setstatus("Pre Analysis Rules")
		LOCAL lni,lcFunc
		FOR lni = 1 TO ALEN(THIS.aRules,1)
			IF EMPTY(THIS.aRules(lni,1))
				LOOP
			ENDIF
			IF ALEN(THIS.aRules,2)>3
				THIS.cWarningID = THIS.aRules(lni,4)
			ELSE
				THIS.cWarningID = ""

			ENDIF
			IF THIS.aRules(lni,2)="I"
				lcRule = THIS.aRules(lni,1)
				lcFunc = THIS.aRules(lni,3)
				TRY
					EXECSCRIPT(lcFunc)
				CATCH TO loErr
					THIS.AddError(loErr,lcRule,lcFunc)
					THIS.aRules(lni,1) = ""
				ENDTRY
			ENDIF
		ENDFOR


	PROCEDURE PostValidate

		THIS.oTherm.setstatus("Post Analysis Rules")
		LOCAL lni,lcFunc
		FOR lni = 1 TO ALEN(THIS.aRules,1)
			IF EMPTY(THIS.aRules(lni,1))
				LOOP
			ENDIF
			IF ALEN(THIS.aRules,2)>3
				THIS.cWarningID = THIS.aRules(lni,4)
			ELSE
				THIS.cWarningID = ""

			ENDIF
			IF THIS.aRules(lni,2)="P"
				lcRule = THIS.aRules(lni,1)
				lcFunc = THIS.aRules(lni,3)
				TRY
					EXECSCRIPT(lcFunc)
				CATCH TO loErr
					THIS.AddError(loErr,lcRule,lcFunc)
					THIS.aRules(lni,1) = ""
				ENDTRY
			ENDIF
		ENDFOR


	PROCEDURE AddError
		LPARAMETERS toErr,tcRule,tcFunc
		THIS.cError = THIS.cError + loErr.MESSAGE+" occurred on line "+LTRIM(STR(loErr.LINENO))+" ("+loErr.LINECONTENTS+") in rule "+tcRule + CHR(13)+CHR(10)

	PROCEDURE Add2Array
		LPARAMETERS tcCode,tnLines,taArray

		IF NOT EMPTY(THIS.aCode(1,1))
			DIMENSION THIS.aCode(ALEN(THIS.aCode,1)+1,4)
		ENDIF

		THIS.cFuncName = tcCode
		IF NOT EMPTY(THIS.cObject)
			THIS.cFuncName = LOWER(THIS.cObject+"."+THIS.cFuncName)
		ELSE
			IF EMPTY(tcCode) OR tcCode="Program"
				THIS.cFuncName = THIS.cFile
			ENDIF
		ENDIF

		THIS.aCode(ALEN(THIS.aCode,1),1) = THIS.cFuncName
		THIS.aCode(ALEN(THIS.aCode,1),2) = tnLines
		THIS.aCode(ALEN(THIS.aCode,1),4) = THIS.cClassName
		IF EMPTY(THIS.aCode(ALEN(THIS.aCode,1),1))
			THIS.aCode(ALEN(THIS.aCode,1),1) = ""
		ENDIF
		LOCAL lcCode
		lcCode = ""
		LOCAL lni
		lnReal=0
		FOR lni = 1 TO ALEN(taArray,1)
			IF THIS.lLineRules
				IF ALLTRIM(STRTRAN(taArray(lni),"	"))<>"*"
					IF NOT EMPTY(taArray(lni))
						THIS.nLine = lni
						THIS.AddToCursor(THIS.cFile,THIS.cFuncName,THIS.cClassName,"Function")
						THIS.ValidateLine(taArray(lni))

						lnReal = lnReal+1
					ENDIF
				ENDIF
			ENDIF
			lcCode = lcCode + taArray(lni)+CHR(13)+CHR(10)
		ENDFOR

		THIS.aCode(ALEN(THIS.aCode,1),3) = lnReal
		THIS.nFuncLines = tnLines
		THIS.AddToCursor(THIS.cFile,THIS.cFuncName,THIS.cClassName,"Function")
		THIS.ValidateCode(lcCode,tcCode)
	ENDPROC

ENDDEFINE

DEFINE CLASS frmResults AS FORM


	DOCREATE = .T.
	AUTOCENTER = .T.
	CAPTION = "Refactoring Results"
	WINDOWTYPE = 1
	WIDTH = 400
	NAME = "Form1"


	ADD OBJECT list1 AS LISTBOX WITH ;
		COLUMNCOUNT = 3, ;
		COLUMNWIDTHS = "250,50,50", ;
		HEIGHT = 144, ;
		LEFT = 24, ;
		TOP = 60, ;
		FONTSIZE=8,;
		WIDTH = 374, ;
		NAME = "List1"


	ADD OBJECT command1 AS COMMANDBUTTON WITH ;
		TOP = 216, ;
		LEFT = 264, ;
		HEIGHT = 27, ;
		WIDTH = 84, ;
		CAPTION = "\<OK", ;
		NAME = "Command1"


	ADD OBJECT label1 AS LABEL WITH ;
		AUTOSIZE = .T., ;
		CAPTION = "Object Name", ;
		HEIGHT = 17, ;
		LEFT = 24, ;
		TOP = 12, ;
		WIDTH = 74, ;
		NAME = "Label1"


	ADD OBJECT label2 AS LABEL WITH ;
		AUTOSIZE = .T., ;
		CAPTION = "Method", ;
		HEIGHT = 17, ;
		LEFT = 24, ;
		TOP = 36, ;
		WIDTH = 42, ;
		NAME = "Label2"


	ADD OBJECT label3 AS LABEL WITH ;
		AUTOSIZE = .T., ;
		CAPTION = "# Lines", ;
		HEIGHT = 17, ;
		LEFT = 228, ;
		TOP = 36, ;
		WIDTH = 43, ;
		NAME = "Label3"


	ADD OBJECT label4 AS LABEL WITH ;
		AUTOSIZE = .T., ;
		CAPTION = "# Code", ;
		HEIGHT = 17, ;
		LEFT = 288, ;
		TOP = 36, ;
		WIDTH = 42, ;
		NAME = "Label4"


	PROCEDURE INIT
		LPARAMETERS tcObj, taArray,tlWork

		IF EMPTY(tlWork)
			THIS.CAPTION = "Object Overview"
		ELSE
			THIS.CAPTION = "Recommended Review Areas"
		ENDIF

		THIS.label1.CAPTION = tcObj

		LOCAL lni,llTitle
		llTitle = .F.
		FOR lni =1 TO ALEN(_SCREEN._Analyst.aWarnings,1)
			WITH THIS.list1
				IF NOT EMPTY(_SCREEN._Analyst.aWarnings(lni,1))
					IF NOT llTitle
						.ADDITEM("*** Warnings ***")
						llTitle = .T.
					ENDIF
					.ADDITEM(_SCREEN._Analyst.aWarnings(lni,1))
				ENDIF
			ENDWITH
		ENDFOR
		THIS.list1.COLUMNCOUNT=1

		IF .F.
			LOCAL lni
			FOR lni =1 TO ALEN(taArray,1)
				WITH THIS.list1
					IF NOT EMPTY(taArray(lni,1))
						.ADDITEM(taArray(lni,1))
						.LIST(.LISTCOUNT,2) = LTRIM(STR(taArray(lni,2)))
						.LIST(.LISTCOUNT,3) = LTRIM(STR(taArray(lni,3)))
					ENDIF
				ENDWITH
			ENDFOR
		ENDIF
		THIS.list1.LISTINDEX = 1
	ENDPROC


	PROCEDURE command1.CLICK
		THISFORM.RELEASE()
	ENDPROC



ENDDEFINE
