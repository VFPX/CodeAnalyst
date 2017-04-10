* Abstract:
*   Class for add/retrieving values
*	from FoxUser resource file.
*
#include "foxref.h"

CLEAR
x = NEWOBJECT("CFoxResource", "FoxRefResource.prg")
*!*	x.Add("Width", 100)
*!*	x.Add("Height", 200)
*!*	x.Add("Name", "Ryan")
x.Load("TEST")

? x.get("Height")
? x.get("Name")
* ? x.save("TEST")

RETURN

DEFINE CLASS CFoxResource AS Session
	PROTECTED oCollection
	
	oCollection  = .NULL.

	ResourceType = "PREFW"
	ResourceFile = ''
	
	PROCEDURE Init()
		SET TALK OFF
		SET DELETED ON
		THIS.oCollection = NEWOBJECT("CCollection", "FoxRefCollection.prg")
		THIS.ResourceFile = SYS(2005)
	ENDPROC

	* Clear out all options
	FUNCTION Clear()
		THIS.oCollection.Remove(-1)
	ENDFUNC
	
	FUNCTION Add(cOption, xValue)
		RETURN THIS.oCollection.Add(m.xValue, UPPER(m.cOption))
	ENDFUNC
	
	FUNCTION Get(cOption)
		RETURN THIS.oCollection.Item(UPPER(m.cOption))
	ENDFUNC
	
	FUNCTION OpenResource()
		IF USED("ResourceAlias")
			USE IN ResourceAlias
		ENDIF

		IF FILE(THIS.ResourceFile)
			TRY
				USE (THIS.ResourceFile) ALIAS FoxResource IN 0 SHARED AGAIN
			CATCH
			ENDTRY
		ENDIF
		
		RETURN USED("FoxResource")
	ENDFUNC
	
	PROCEDURE Save(cID, cName)
		LOCAL nSelect
		LOCAL cType
		LOCAL i
		LOCAL ARRAY aOptions[1]
		
		IF VARTYPE(m.cName) <> 'C'
			m.cName = ''
		ENDIF
		
		IF THIS.OpenResource()
			m.nSelect = SELECT()
			
			m.cType = PADR(THIS.ResourceType, LEN(FoxResource.Type))
			m.cID   = PADR(m.cID, LEN(FoxResource.ID))

			SELECT FoxResource
			LOCATE FOR Type == m.cType AND ID == m.cID AND Name == m.cName
			IF !FOUND()
				APPEND BLANK IN FoxResource
				REPLACE ; 
				  Type WITH m.cType, ;
				  ID WITH m.cID, ;
				  ReadOnly WITH .F. ;
				 IN FoxResource
			ENDIF

			IF !FoxResource.ReadOnly
				IF THIS.oCollection.Count > 0
					DIMENSION aOptions[THIS.oCollection.Count, 2]
					FOR m.i = 1 TO THIS.oCollection.Count
						aOptions[m.i, 1] = THIS.oCollection.GetKey(m.i)
						aOptions[m.i, 2] = THIS.oCollection.Item(m.i)
					ENDFOR
					SAVE TO MEMO Data ALL LIKE aOptions
				ELSE
					BLANK FIELDS Data IN FoxResource
				ENDIF

				REPLACE ;
				  Updated WITH DATE(), ;
				  ckval WITH VAL(SYS(2007, FoxResource.Data)) ;
				 IN FoxResource
			ENDIF
			
			SELECT (m.nSelect)
		ENDIF
	ENDPROC
	
	PROCEDURE Load(cID, cName)
		LOCAL nSelect
		LOCAL cType
		LOCAL ARRAY aOptions[1]
		
		IF VARTYPE(m.cName) <> 'C'
			m.cName = ''
		ENDIF
		
		THIS.Clear()
		IF THIS.OpenResource()
			m.nSelect = SELECT()
			
			m.cType = PADR(THIS.ResourceType, LEN(FoxResource.Type))
			m.cID   = PADR(m.cID, LEN(FoxResource.ID))

			SELECT FoxResource
			LOCATE FOR Type == m.cType AND ID == m.cID AND Name == m.cName
			IF FOUND() AND !EMPTY(Data) AND ckval == VAL(SYS(2007, Data))
				RESTORE FROM MEMO Data ADDITIVE
				IF VARTYPE(aOptions[1,1]) == 'C'
					m.nCnt = ALEN(aOptions, 1)
					FOR m.i = 1 TO m.nCnt
						THIS.Add(aOptions[m.i, 1], aOptions[m.i, 2])
					ENDFOR
				ENDIF
			ENDIF
			
			SELECT (m.nSelect)
		ENDIF
	ENDPROC

ENDDEFINE

