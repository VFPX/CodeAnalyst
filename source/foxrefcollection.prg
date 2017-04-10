* Abstract:
* 	Collection Class.  
*	This could be replaced with a subclass of
*	the VFP8 collection class, which wasn't
*	available when Code References work began.
* 	
DEFINE CLASS CFoxRefCollection AS Custom
	DIMENSION aCollection[1, 2]
	Count = 0
	
	FUNCTION Add(xValue, cKey)
		THIS.Count = THIS.Count + 1
		DIMENSION THIS.aCollection[THIS.Count, 2]

		THIS.aCollection[THIS.Count, 1] = m.xValue
		IF VARTYPE(m.cKey) == 'C'
			THIS.aCollection[THIS.Count, 2] = m.cKey
		ELSE
			THIS.aCollection[THIS.Count, 2] = ''
		ENDIF
		
		RETURN .T.
	ENDFUNC
	
	FUNCTION Remove(nIndex)
		IF m.nIndex == -1  && remove all
			DIMENSION THIS.aCollection[1, 2]
			THIS.Count = 0
			RETURN .T.
		ENDIF

		IF BETWEEN(m.nIndex, 1, THIS.Count)
			=ADEL(THIS.aCollection, m.nIndex, 2)
			THIS.Count = THIS.Count - 1
			RETURN .T.
		ELSE
			RETURN .F.
		ENDIF
	ENDFUNC

	FUNCTION Item(xIndex)
		LOCAL i
		LOCAL nCnt

		DO CASE
		CASE VARTYPE(m.xIndex) == 'N'
			IF BETWEEN(m.xIndex, 1, THIS.Count)
				RETURN THIS.aCollection[m.xIndex, 1]
			ELSE
				RETURN .NULL.
			ENDIF

		CASE VARTYPE(m.xIndex) == 'C'
			m.nCnt = THIS.Count
			FOR m.i = 1 TO m.nCnt
				IF THIS.aCollection[m.i, 2] == m.xIndex
					RETURN THIS.aCollection[m.i, 1]
				ENDIF
			ENDFOR

		ENDCASE
		
		RETURN .NULL.
	ENDFUNC

	* given an index, return the key
	* or given a Key, return the index
	FUNCTION GetKey(xIndex)
		LOCAL nCnt
		LOCAL i

		DO CASE
		CASE VARTYPE(m.xIndex) == 'N'  && look for index, return key
			IF BETWEEN(m.xIndex, 1, THIS.Count)
				RETURN THIS.aCollection[m.xIndex, 2]
			ELSE
				RETURN ''
			ENDIF

		CASE VARTYPE(m.xIndex) == 'C'  && look for key, return index
			m.nCnt = THIS.Count
			FOR m.i = 1 TO m.nCnt
				IF THIS.aCollection[m.i, 2] == m.xIndex
					RETURN m.i
				ENDIF
			ENDFOR
			
			RETURN -1
		ENDCASE
		
		RETURN .NULL.
	ENDFUNC


	FUNCTION SetItem(nIndex, xValue)
		IF BETWEEN(m.nIndex, 1, THIS.Count)
			THIS.aCollection[m.nIndex, 1] = m.xValue
			RETURN .T.
		ELSE
			RETURN .F.
		ENDIF
	ENDFUNC


	FUNCTION AddNoDupe(xValue)
		LOCAL i
		LOCAL nCnt
		LOCAL lFound
		
		m.lFound = .F.
		m.nCnt = THIS.Count
		FOR m.i = 1 TO m.nCnt
			IF THIS.aCollection[m.i, 1] == xValue
				m.lFound = .T.
				EXIT
			ENDIF
		ENDFOR
		IF !m.lFound
			THIS.Add(xValue)
		ENDIF
	ENDFUNC

	* clear collection
	FUNCTION Clear()
		RETURN THIS.Remove(-1)
	ENDFUNC

	* add a delimited list to the collection	
	FUNCTION AddList(cList, cParseChar)
		LOCAL nCnt
		LOCAL i
		LOCAL ARRAY aList[1]
		
		IF VARTYPE(m.cParseChar) <> 'C' OR LEN(m.cParseChar) == 0
			m.cParseChar = ','
		ENDIF
		m.nCnt = ALINES(m.aList, m.cList, .F., m.cParseChar)
		FOR i = 1 TO m.nCnt
			IF !EMPTY(m.aList[m.i])
				THIS.Add(m.aList[m.i])
			ENDIF
		ENDFOR

		RETURN m.nCnt		
	ENDFUNC
	
	* given a value, return the index position within the collection
	FUNCTION GetIndex(xValue)
		LOCAL nCnt
		LOCAL i
		LOCAL nIndex
		
		m.nIndex = 0
		m.nCnt = THIS.Count
		FOR m.i = 1 TO m.nCnt
			IF THIS.Item(m.i) == xValue
				m.nIndex = m.i
				EXIT
			ENDIF
		ENDFOR
		
		RETURN m.nIndex
	ENDFUNC
ENDDEFINE
