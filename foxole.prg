*-- Programa: FoxOle (Componente para manipular recordsets de ADO y cursor de Fox)
*-- Fecha: 19/04/2002
*-- Autor: Angel Jesús Hernández
*-- E-mail: angeljesus14v@hotmail.com
*-- Lenguaje: Microsoft Visual FoxPro 7 SP1
*-- Notas: Se tuvo que crear este componente por problemas al usar el componente proporcionado por Microsoft
*--        (vfpcom), ya que este tiene "bugs" a la hora de convertir de un tipo de datos a otro. El primer VFPCOM convierte
*--        sin problemas los tipos TinyInt de SQL pero no funciona con los tipos flotantes, en el segundo VFPCOM sucede
*--        a la inversa, es decir, los flotantes se convierten sin problemas pero los TinyInt los convierte a campos
*--        generales. Este componente se llama diferente  al proporcionado por Microsoft para respetar el derecho de autor.
*--

#Define True .T.
#Define False .F.
#INCLUDE AdoVfp.h  && Archivo de cabecera con definiciones de constantes para ADO

Define Class CursorUtil As Session OlePublic
	Protected oRst As Object
	oRst=Null && Recordset de trabajo
	nHwnd=0  && Handle de la ventana principal de la aplicación
	oApp = Null
	*-- Constructor
	Procedure Init() HelpString "Constructor del COMponente"
		*-- Opciones del entorno
		Set Talk Off
		Set Safety Off
		Set StrictDate To 0
		*-- Inicializamos los objetos usados por el componente
		With This
			.oRst=Createobject("adodb.recordset")
			*-- En caso que no se encuentre instalado MDAC
			If Type(".oRst") # "O" And !Isnull(.oRst)
				.oApp.Application.DoCmd("MsgBox("+Alltrim(Str(.nHwnd))+;
				",'No se pudo crear objeto de OLEDB (Recordset). Verifique que ha instalado el MDAC','Información',0x00000030)")
				Return False
			Endif
		Endwith
	Endproc
	*-- Destructor
	Procedure Destroy() HelpString "Destructor del COMponente"
		*-- Guardamos los errores ocurridos durante la ejecución del componente
		Local lcRuta As String
		lcRuta=Addbs(Getenv("temp"))+"FoxOle_log.xml"
		*-- Cerramos el cursor encargado de guardar los errores ocurridos en el componente
		If This.oApp.Application.Eval("USED('FoxOle_log')")
			If This.oApp.Application.Eval("!EMPTY(RECCOUNT('FoxOle_log'))")
				=This.oApp.Application.DoCmd("CURSORTOXML('FoxOle_log', ' "+lcRuta+" ',1, 512, 0)")
			Endif
			=This.oApp.Application.DoCmd("USE IN FoxOle_log")
		Endif
		*-- Liberamos los objetos usados
		Store Null To This.oRst, This.oApp
	Endproc
	*-- Control de errores
	Procedure Error(nError, cMethod, nLine) As void HelpString "Control de Errores"
		If This.oApp.Application.Eval("USED('FoxOle_log')")
			This.oApp.Application.DoCmd("INSERT INTO FoxOle_log (dError, cMethod, cMensaje, cCodigo, nSesion) VALUES "+;
			"(DATETIME(), PROGRAM(),MESSAGE(),MESSAGE(1),SET('DATASESSION'))")
		Endif
	Endproc
	*-- Convierte de Recordset a Cursor
	Function RsToCursor(oRecordset As Object, cCursor As String) As Boolean ;
		helpstring "Convierte de recordset a Cursor"
		Local lcRst As String,;
		lcTmpFile As String,;
		lcTmpDir As String,;
		lcRuta As String,;
		lcGen As String,;
		lnIndice As Integer,;
		lGenField As Boolean,;
		lUsado As Boolean

		Local objCampos As Object,;
		laInfoRst As Array,;
		lcPoint As String,;
		lcSeparator As String,;
		lnBufer As Integer,;
		lMemoField As Boolean,;
		lcMemo As String,;
		lcExpr As String,;
		laDirTemp[1] As Array

		Store "" To lcRst, lcTmpFile, lcTmpDir, lcRuta, lcDecimal, lcGen,;
		lcPoint, lcSeparator, lcMemo, lcExpr
		Store 0 To lnIndice, lnBufer
		lcTmpFile=Sys(2015) && Nombre del archivo que recibirá la información del recordset
		lcTmpDir=Addbs(Getenv("TEMP")) && Directorio temporal
		*-- Verificamos los argumentos pasados al método
		If PCOUNT() # 2 Or Type("oRecordset") # "O" Or Type("oRecordset.fields(0)") # "O" Or;
			ISNULL(oRecordset.Fields(0)) Or Type("cCursor") # "C"
			=This.oApp.Application.DoCmd("MsgBox("+Alltrim(Str(This.nHwnd))+;
			",'Los parámetros pasados al método no son correctos. Verifique por favor','Error en método RsToCursor',0x00000030)")
			Return False
		Endif
		*-- Convertimos de recordset a cursor
		With oRecordset
			Dimension laInfoRst[.Fields.Count,16]
			For Each objCampo In .Fields
				lnIndice=lnIndice+1
				*-- Elementos no utilizados para la creación del cursor (reglas, desencadenantes, etc) si fuese el caso de un
				*-- Dataset de ADO .Net si sería posible pero el ADO convencional no trae consigo estas propiedades ya que
				*-- son evaluadas a nivel del servidor
				Store "" To laInfoRst[lnIndice,7],laInfoRst[lnIndice,8],laInfoRst[lnIndice,9],;
				laInfoRst[lnIndice,10],laInfoRst[lnIndice,11],laInfoRst[lnIndice,12],;
				laInfoRst[lnIndice,13],laInfoRst[lnIndice,14],laInfoRst[lnIndice,15],;
				laInfoRst[lnIndice,16]
				laInfoRst[lnIndice,1]=objCampo.Name && Nombre del campo
				laInfoRst[lnIndice,3]=Iif(!Inlist(objCampo.Type,ADBINARY,ADLONGVARBINARY,ADVARBINARY), objCampo.DefinedSize,0) && Ancho del campo
				laInfoRst[lnIndice,4]=Iif(Inlist(objCampo.Type,ADDOUBLE,ADCURRENCY,ADDECIMAL,ADNUMERIC,; && Precisión del campo (decimales)
				ADERROR,ADUSERDEFINED,ADIDISPATCH,ADIUNKNOWN,ADGUID,ADSINGLE),This.oApp.Application.Eval("Set('decimals')"),0)
				laInfoRst[lnIndice,5]=Bittest(objCampo.Attributes,5) && ¿Permite valores nulos?
				laInfoRst[lnIndice,6]=False && ¿Permite cambiar página de códigos?
				*-- Determinamos el tipo de campo que almacenará el campo
				Do Case
					Case Inlist(objCampo.Type,ADTINYINT,ADSMALLINT,ADINTEGER,ADBIGINT,ADUNSIGNEDTINYINT,;
						ADUNSIGNEDSMALLINT,ADUNSIGNEDINT,ADUNSIGNEDBIGINT,ADBOOLEAN) && Entero/Booleano
						laInfoRst[lnIndice,2]="I"
					Case Inlist(objCampo.Type,ADBSTR,ADCHAR,ADVARCHAR,ADWCHAR,ADVARWCHAR,ADLONGVARCHAR) && Caracter
						laInfoRst[lnIndice,2]="C"
						laInfoRst[lnIndice,3]=Iif(objCampo.DefinedSize <= 254, objCampo.DefinedSize,254)
					Case Inlist(objCampo.Type,ADBINARY,ADLONGVARBINARY,ADVARBINARY) && General (Binario)
						laInfoRst[lnIndice,2]="G"
						lGenField=True
						lcGen=objCampo.Name
					Case Inlist(objCampo.Type,ADCURRENCY) && Moneda
						laInfoRst[lnIndice,2]="Y"
					Case Inlist(objCampo.Type, ADDATE,ADDBDATE) && Fecha
						laInfoRst[lnIndice,2]="D"
					Case Inlist(objCampo.Type,ADDBTIME,ADDBTIMESTAMP) && Fecha/hora
						laInfoRst[lnIndice,2]="T"
					Case Inlist(objCampo.Type,ADNUMERIC,ADERROR,ADIUNKNOWN,ADIDISPATCH,ADGUID) && Númerico
						laInfoRst[lnIndice,2]="N"
					Case Inlist(objCampo.Type,ADDECIMAL,ADDOUBLE,ADSINGLE) && Decimal/flotante
						laInfoRst[lnIndice,2]="B"
					Case Inlist(objCampo.Type,ADLONGVARWCHAR) && Memo
						laInfoRst[lnIndice,2]="M"
						lMemoField=True
						lcMemo=objCampo.Name
				Endcase
			Endfor
		Endwith
		*-- Creamos el cursor y variables del lado de la aplicación (por medio del objeto APPLICATION pasado al COM). Verificamos si
		*-- existe un alias igual al especificado por el parámetro cCursor, si es así "zapeamos" dicho alias. En caso contrario creamos
		*-- el cursor basándonos en la información contenida en la matriz laInfoRst
		This.oApp.Application.DoCmd("SET SAFETY OFF")
		lUsado=This.oApp.Application.Eval("USED('"+cCursor+"')")
		lcPoint=This.oApp.Application.Eval("SET('POINT')")
		lcSeparator=This.oApp.Application.Eval("SET('SEPARATOR')")
		With This
			If Not lUsado
				.oApp.Application.SetVar("ajhmgr", @laInfoRst)
				.oApp.Application.SetVar("mgrajh", cCursor)
				.oApp.Application.DoCmd("CREATE CURSOR (mgrajh) FROM ARRAY ajhmgr")
			Else
				*-- El comando ZAP solo se puede ejecutar con cursores que tengan el modo de almacenamiento en búferes de fila (3)
				*-- Por ello cambiamos el modo para "zapear" el cursor y después lo restauramos.
				lnBufer=This.oApp.Application.Eval("CURSORGETPROP('Buffering','"+cCursor+"')")
				If This.oApp.Application.Eval("CURSORGETPROP('Buffering','"+cCursor+"')") # 1
					=This.oApp.Application.DoCmd("TABLEREVERT(.T.)")
					=This.oApp.Application.DoCmd("CURSORSETPROP('Buffering',1,'"+cCursor+"')")
				Endif
				= This.oApp.Application.DoCmd("ZAP IN "+ cCursor)
			Endif
		Endwith
		*-- Verificamos si el cursor se creó existosamente
		If Not This.oApp.Application.Eval("USED('"+cCursor+"')")
			Return False
		Endif
		*-- Guardamos la información del recordset en un archivo de texto temporal
		With This.oApp.Application
			If !Empty(oRecordset.recordcount)    && Copiamos la información del recordset al cursor si el primero contiene data
				=oRecordset.MoveFirst()
				.DoCmd("SET STRICTDATE TO 0")
				.DoCmd("SET CENTURY ON")
				.DoCmd("SET DATE TO BRITISH")
				.DoCmd("SET SEPARATOR TO '.'")
				lcRst=oRecordset.GetString(2,oRecordset.recordcount,Chr(9),Chr(13),'.NULL.')
				lcRuta=lcTmpDir+lcTmpFile+".dat"
				*-- Revisamos si existen campos fecha/hora y entonces reemplazamos el a.m. por am y p.m. por pm esto por la sencilla razón
				*-- de que estos tipos de campos tienden a causar conflictos con Fox y al momento de agregarlos al cursor estos son agregados
				*-- en blanco
				lcRst=Strtran(lcRst, "a.m.","am")  && Buscamos a.m.
				lcRst=Strtran(lcRst, "p.m.","pm") && Buscamos p.m.
				=Strtofile(lcRst,lcRuta)
				*-- Agregamos la información del archivo al cursor recién creado
				.DoCmd("SET POINT TO ','")
				.DoCmd("APPEND FROM '"+lcRuta+"' DELIMITED WITH TAB")
				.DoCmd("CURSORSETPROP('Buffering',"+Alltrim(Str(lnBufer))+", '"+ cCursor+"')")
				*-- Si existe al menos un campo general procedemos a agregarlo al cursor
				If lGenField
					=This.ObtenCampoGeneral(cCursor, oRecordset, lcGen)
				Endif
				*-- Si existe al menos un campo memo procedemos a agregarlo al cursor
				If lMemoField
					=This.ObtenCampoMemo(cCursor, oRecordset, lcMemo)
				Endif
				*-- Liberamos variables y eliminamos archivos temporales
				lcExpr=Addbs(Sys(2023))+"*.atm"
				.DoCmd("RELEASE ajhmgr, mgrajh")
				.DoCmd("SET POINT TO '"+lcPoint+"'")
				.DoCmd("SET SEPARATOR TO '"+lcSeparator+"'")
				Delete File "&lcRuta"
				*-- Eliminamos el archivo que contiene información de los campos memo
				If !Empty(Adir(laDirTemp,lcExpr))
					Delete File "&lcExpr"
				Endif
			Endif
		Endwith
		Return True
	Endfunc
	*-- Agrega un campo memo traído desde un campo texto
	Hidden Function ObtenCampoMemo(cCursor As String, oRecordset As Object,;
		lcMemo As String) As void HelpString "Agrega un campo memo traído desde un campo texto"
		Local lcAlias As String,;
		lcExpr As String,;
		lcTmpFile As String
		Store "" To lcAlias, lcExpr, lcTmpFile

		lcAlias=This.oApp.Application.Eval("ALIAS()")
		=This.oApp.Application.DoCmd("SELECT "+cCursor)
		=This.oApp.Application.DoCmd("GO TOP")
		=oRecordset.MoveFirst()

		Do While Not oRecordset.Eof
			lcTmpFile=Addbs(Sys(2023))+Sys(2015)+".atm"
			=Strtofile(oRecordset.Fields("&lcMemo").Value, lcTmpFile)
			lcExpr = "Append Memo "+cCursor+"."+lcMemo+" From '"+lcTmpFile+"'"
			=This.oApp.Application.DoCmd(lcExpr)
			=oRecordset.MoveNext()
			=This.oApp.Application.DoCmd("SKIP")
		Enddo

		If !Empty(lcAlias)
			=This.oApp.Application.DoCmd("SELECT "+lcAlias)
		Endif
	Endfunc
	*-- Agrega un campo general traído desde un campo binario
	Hidden Function ObtenCampoGeneral(cCursor As String, oRecordset As Object,;
		lcGen As String) As void HelpString "Agrega un campo general traído desde un campo binario"
		Local lcTmpBmp As String,;
		lcExpr As String,;
		lcAlias As String,;
		lcStrBmp As String
		Store "" To lcExpr, lcStrBmp
		lcAlias=This.oApp.Application.Eval("ALIAS()")
		*-- Extraemos la información de la imagen y la guardamos en un archivo temporal para agregarla luego al campo
		*-- general. Inicialmente procesaremos solo mapa de bits (BMP) después que hayamos analizado los demás
		*-- encabezados de los archivos de imagenes lo iremos incorporando luego...
		=This.oApp.Application.DoCmd("SELECT "+cCursor)
		=This.oApp.Application.DoCmd("GO TOP")
		=oRecordset.MoveFirst()
		*-- Recorremos el recordset y solo procesamos los campos generales
		Do While Not oRecordset.Eof
			lcStrBmp=Alltrim(oRecordset.Fields("&lcGen").Value)
			lcExpr=Iif(!Empty(Atc("BM",lcStrBmp)),Substr(lcStrBmp,Atc("BM",lcStrBmp)),"")
			If !Empty(lcExpr)
				lcTmpBmp=Sys(2015)+".bmp" && Archivo temporal
				=Strtofile(lcExpr,"&lcTmpBmp")
				*-- Agregamos la imagen desde el archivo temporal
				=This.oApp.Application.DoCmd("APPEND GENERAL "+lcGen+" FROM "+lcTmpBmp+" CLASS PAINT.PICTURE")
				*-- Eliminamos el archivo temporal
				Delete File (lcTmpBmp)
			Endif
			=oRecordset.MoveNext()
			=This.oApp.Application.DoCmd("SKIP")
		Enddo
		If !Empty(lcAlias)
			=This.oApp.Application.DoCmd("SELECT "+lcAlias)
		Endif
	Endfunc
	*-- Asignamos el objeto pasado como parámetro a la propiedad oApp
	Function init2 (objApp As Object) As void HelpString "Inicializador del COMponente"
		With This
			.oApp=objApp
			.nHwnd=oApp.HWnd
			*-- Cursor encargado de almacenar los errores
			.oApp.Application.DoCmd("CREATE CURSOR FoxOle_Log (dError T, cMethod c(50), cMensaje c(50),cCodigo c(50), nSesion N(4,0))")
			*-- Funciones del API a utilizar
			.oApp.Application.DoCmd("DECLARE MessageBox IN WIN32API AS MsgBox INTEGER, STRING, STRING, SHORT")
		Endwith
	Endfunc
	*-- Información del componente
	Function InfoCOM As void HelpString "Información del COMponente"
		=This.oApp.Application.DoCmd("MsgBox("+Alltrim(Str(This.nHwnd))+;
		",'Utilidad para convertir recordsets de ADO a cursores de Visual FoxPro','Acerca de FoxOle',0)")
	Endfunc
	*-- Convierte un cursor de FoxPro en un recordset de ADO
	Function CursorToRs(lcAlias As String,lXMLSource As Boolean, lNoData As Boolean) As ADODB.Recordset;
		helpstring "Convierte un cursor de FoxPro en un recordset de ADO"

		Local lcOldAlias As String,;
		luValor As String,;
		laInfoFld As Array,;
		lcCampo As String,;
		lnCampo As Integer,;
		loRst As ADODB.Recordset,;
		lcCompa As String
		lcCompa=Set("Compatible")

		Store "" To lcOldAlias, lcCom, luValor, lcCampo
		lcOldAlias=Alias()
		Store 0 To lnCampo, lnPaso

		With This.oApp.Application
			If Type("lcAlias") # "C" Or Empty(lcAlias)
				.DoCmd("MsgBox("+Alltrim(Str(.HWnd))+;
				",'El parámetro lcAlias no es correcto. Verifique por favor','Error en método CursorToRs',1")
				Return .Null.
			Endif
			If !.Eval("USED('"+lcAlias+"')")
				.DoCmd("MsgBox("+Alltrim(Str(.HWnd))+;
				",'El Alias especificado no es válido. Verifique por favor','Error en método CursorToRs',1")
				Return .Null.
			Endif
		Endwith
		*-- Traemos el cursor de la aplicación cliente al COM
		If This.BringCursorToRPC(lcAlias,lXMLSource)
			Select (lcAlias)
			Set Compatible Off
			*-- Creamos el recordset que será devuelto por el método
			loRst=Createobject("adodb.recordset")
			loRst.CursorType= 3  && adOpenStatic
			loRst.LockType= 2  && adLockPessimistic
			*-- Creamos la estructura del recordset (campos, ancho, etc)
			Dimension laInfoFld[FCOUNT()]
			With loRst
				For lnCampo=1 To Fcount()
					laInfoFld[lnCampo]=Type("EVALUATE(FIELD(lnCampo))")
					Do Case
						Case Type("EVALUATE(FIELD(lnCampo))")="C" && caracter
							.Fields.Append(Field(lnCampo),ADVARCHAR,Fsize(Field(lnCampo)))
						Case Type("EVALUATE(FIELD(lnCampo))")="N" && númerico
							luValor=Evaluate(Field(lnCampo))
							luValor=Alltrim(Str(luValor,Len(Alltrim(Str(luValor)))+3,2))
							Do Case
								Case Empty(At(".",luValor)) And Between(Val(luValor),0,255)  And  Between(Len(luValor),1,3)   && TinyInt
									.Fields.Append(Field(lnCampo),ADTINYINT,Fsize(Field(lnCampo)),0)
								Case Empty(At(".",luValor)) And Between(Val(luValor),-32768,32767) && SmallInt
									.Fields.Append(Field(lnCampo),ADSMALLINT,Fsize(Field(lnCampo)),0)
								Case Empty(At(".",luValor)) And Between(Val(luValor),-32768,32767) && Integer
									.Fields.Append(Field(lnCampo),ADINTEGER,Fsize(Field(lnCampo)),0)
								Otherwise && Float
									.Fields.Append(Field(lnCampo),ADDOUBLE,Fsize(Field(lnCampo)),2)
							Endcase
						Case Type("EVALUATE(FIELD(lnCampo))")="Y" && moneda
							.Fields.Append(Field(lnCampo),ADCURRENCY,Fsize(Field(lnCampo)), 63)
						Case Type("EVALUATE(FIELD(lnCampo))")="T" && fecha/hora
							.Fields.Append(Field(lnCampo),ADDBTIMESTAMP,22)
						Case Type("EVALUATE(FIELD(lnCampo))")="L" && lógica
							.Fields.Append(Field(lnCampo),ADDBOOLEAN,1)
						Case Type("EVALUATE(FIELD(lnCampo))")="D" && fecha
							.Fields.Append(Field(lnCampo),ADDDBDATE,10)
						Case Type("EVALUATE(FIELD(lnCampo))")="G" && General
							.Fields.Append(Field(lnCampo),ADLONGVARBINARY,256)
					Endcase
				Endfor
			Endwith
			*-- Procedemos a llenar el recordset con la información traída del cursor (verificamos que no hayan pasado .T.
			*-- en el parámetro lNoData y que el cursor no este vacío)
			Go Top In (lcAlias)
			If !lNoData And !Empty(Reccount(lcAlias))
				With loRst
					.Open()
					Scan
						.AddNew() && Agregamos un nuevo registro al recordset
						For lnCampo = 1 To Fcount()
							lcCampo=Field(lnCampo)
							Do Case
								Case laInfoFld[lnCampo] = "C"
									.Fields(lcCampo).Value=Evaluate(Field(lnCampo))
								Case laInfoFld[lnCampo] = "N" Or laInfoFld[lnCampo] = "Y"
									.Fields(lcCampo).Value=Evaluate(Field(lnCampo))
								Case laInfoFld[lnCampo] = "T"
									.Fields(lcCampo).Value=Alltrim(Ttoc(Evaluate(Field(lnCampo))))
								Case laInfoFld[m.campo] = "L"
									If Evaluate(Field(lnCampo))=True Or Evaluate(Field(lnCampo))=1
										.Fields(lcCampo).Value=1
									Else
										.Fields(lcCampo).Value=0
									Endif
								Case laInfoFld[lnCampo] = "D"
									.Fields(lcCampo).Value=Alltrim(Dtoc(Evaluate(Field(lnCampo))))
								Case laInfoFld[lnCampo] = "G"
									*-- Escribir aquí el código para guardar el campo general
							Endcase
						Endfor
					Endscan
				Endwith
			Endif
		Else
			Return .Null.
		Endif
		*--
		If !Empty(lcOldAlias)
			Select (lcOldAlias)
		Endif
		Set Compatible &lcCompa
		Return loRst && Devolvemos el recordset recién creado
	Endfunc
	*-- Trae un cursor del lado de la aplicación que invoca al componente al lado de este (RPC)
	Hidden Function BringCursorToRPC(lcAlias As String, lXMLSource As Boolean) As Boolean
		Local lcDBFPath As String,;
		lcCurAlias As String,;
		lcCOMAlias As String,;
		lcMacro As String

		lcCOMAlias = Alias()
		lcDBFPath=Addbs(Sys(2023))+Sys(2015)+".tmp"

		With This.oApp.Application
			If Type("lcAlias") # "C" Or Empty(lcAlias)
				.DoCmd("MsgBox("+Alltrim(Str(.HWnd))+;
				",'El parámetro lcAlias no es correcto. Verifique por favor','Error en método BringCursorToRPC',1")
				Return False
			Endif
			If !.Eval("USED('"+lcAlias+"')")
				.DoCmd("MsgBox("+Alltrim(Str(.HWnd))+;
				",'El alias especificado no existe. Verifique por favor','Error en método BringCursorToRPC',1")
				Return False
			Endif
			*-- Volcamos al disco el contenido del cursor para así manejarlo desde el espacio de memoria del COM sin
			*-- tener que hacer referencia al mismo a través del objeto application (_VFP)
			lcCurAlias=.Eval("Alias()")
			.DoCmd("Select '"+lcAlias+"'")
			*-- Creamos el cursor a partir de una tabla temporal ó un archivo XML
			If !lXMLSource
				.DoCmd("Copy To '"+lcDBFPath+"'")
			Else
				.DoCmd("CursorToXML('"+lcAlias+"','"+lcDBFPath+"',1,512,0,'')")
			Endif
			*-- Seleccionamos el área de trabajo original (Aplicación)
			If !Empty(lcCurAlias)
				.DoCmd("Select '"+lcCurAlias+"'")
			Endif
		Endwith
		*-- Verificamos si el haber volcado en el disco el cursor fue existoso
		If !File(lcDBFPath)
			Return False
		Else
			If !lXMLSource
				*-- Creamos el cursor del lado del COM, eliminamos la tabla temporal desde la cual se crea
				lcMacro="Select * from "+Juststem(lcDBFPath)+" into cursor "+lcAlias
				Use (lcDBFPath) In 0 Exclusive
				&lcMacro
				lcMacro="Use In '"+Juststem(lcDBFPath)+"'"
				&lcMacro
				lcMacro="Delete File '"+lcDBFPath+"'"
				&lcMacro
			Else
				=Xmltocursor(lcDBFPath,lcAlias,512)
				Delete File (lcDBFPath)
			Endif
			*-- Seleccionamos el área de trabajo original (COM)
			If !Empty(lcCOMAlias)
				Select (lcCOMAlias)
			Endif
		Endif
		Return True
	Endfunc
	*--
	Function RsToDBF(oRecordset As Object, cRuta As String) As Boolean ;
		helpstring "Guarda un recordset en una tabla de Fox (DBF)"
		Local lcAlias As String,;
		lcCursor As String
		*-- Verificamos los argumentos pasados al método
		With This
			If PCOUNT() # 2 Or Type("oRecordset") # "O" Or Type("oRecordset.fields(0)") # "O" Or;
				ISNULL(oRecordset.Fields(0)) Or Type("cRuta") # "C"
				.oApp.Application.DoCmd("MsgBox("+Alltrim(Str(.nHwnd))+;
				",'Los parámetros pasados al método no son correctos. Verifique por favor','Error en método RsToDBF',0x00000030)")
				Return False
			Endif

			lcCursor=Juststem(cRuta)

			If !.RsToCursor(oRecordset, lcCursor)
				Return False
			Endif

			lcAlias=.oApp.Application.Eval("Alias()")
			.oApp.Application.DoCmd("Select '"+lcCursor+"'")
			.oApp.Application.DoCmd("Copy to '"+cRuta+"' Type Fox2X")
			.oApp.Application.DoCmd("Use in '"+lcCursor+"'")

			If !Empty(lcAlias) And .oApp.Application.Eval("Used('"+lcAlias+"')")
				.oApp.Application.DoCmd("Select '"+lcAlias+"'")
			Endif

			Return True
		Endwith
	Endfunc
Enddefine

