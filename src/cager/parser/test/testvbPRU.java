package cager.parser.test;

public class testvbPRU{

	private Long OpenProcess ( Long dwAccess, Integer fInherit, Long hObject ) {

/* 
 
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwAccess As Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long 
 */
	}

	private Long WaitForSingleObject ( Long hHandle, Long lngMilliseconds ) {

/* 
 Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal lngMilliseconds As Long) As Long 
 */
	}

	private Long CloseHandle ( Long hObject ) {

/* 
 Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long 
 */
	}

	private void Sleep ( Long dwMilliseconds ) {

/* 
 Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 
 */
	}

	private Long GetExitCodeProcess ( Long hProcess, Long lpExitCode ) {

/* 
 Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long 
 */
	}

	private void Run ( String cmd, VBA.VbAppWinStyle WindowStyle, Boolean Wait ) {

/* 
 
Public Function Run(ByVal cmd As String, Optional WindowStyle As VBA.VbAppWinStyle = vbNormalFocus, Optional ByVal Wait As Boolean) As Long
    Debug.Print "RUN: " & cmd
    m_ProgHandle = Shell(cmd, WindowStyle)
    
    If m_ProgHandle = 0 Then
        ' Failed
        Run = -1
    Else
        Run = WaitOnProgram()
        Debug.Print "EXIT " & Run
    End If
    
    If suma (3,4) = 7 Then
    	Run -1
   	Else
   		Run = WaitOnProgram()
        Debug.Print "EXIT " & Run	
   	End If
   	
End Function 
 */
	}

	private void a ( ) {

/* 
 
Public Function a () As Long

   a = 3 + 5

End Function 
 */
	}

	private void Suma ( Long a, Long b ) {

/* 
 


Public Function Suma ( ByVal a As Long, ByVal b As Long ) As Long

	Suma = a + b 
	
	Suma = b + a + Suma (3,4)

End Function 
 */
	}

	private void WaitOnProgram ( ) {

/* 
 
Private Function WaitOnProgram() As Long
    Set m_WaitForm = New frmWait
    m_WaitForm.ShowWait 500
    
    WaitOnProgram = m_ExitCode
End Function 
 */
	}

	private void m_WaitForm_Timer ( ) {

/* 
 
Private Sub m_WaitForm_Timer()
    Dim iExit As Long
    Dim hProg As Long
    
        
    Dim aux As Long
    'Dim number As Integer = 8
    'Dim index As Integer = 0
    Dim Counts As Integer
    Dim found As Boolean 
	Dim thisCollection As New Collection
    Dim counter As Integer
    
    
    hProg = OpenProcess(PROCESS_ALL_ACCESS, False, m_ProgHandle)
    
    GetExitCodeProcess hProg, iExit
    
    CloseHandle hProg
    
    
    'Sentencia IF
    If iExit <> STILL_ACTIVE Then
        m_ExitCode = iExit
        m_WaitForm.Visible = False
        Set m_WaitForm = Nothing
    End If
    
    
    'Bucle FOR
	FOR i=0 TO 6 STEP 2
		aux = i
	NEXT 

	
	'Sentencia CASE
	Select Case number
	Case 1 To 5
	    Debug.WriteLine("Between 1 and 5, inclusive")
	    ' The following is the only Case clause that evaluates to True.
	Case 6, 7, 8
	    Debug.WriteLine("Between 6 and 8, inclusive")
	Case IS = 9, IS = 10
	    Debug.WriteLine("Equal to 9 or 10")
	Case Else
	    Debug.WriteLine("Not between 1 and 10, inclusive")
	End Select

	'Sentencias DO LOOP UNTIL y DO WHILE LOOP
	Check = True ' Se inicializan las variables.
	Counts = 0
	Do ' Empieza sin comprobar ninguna condición.
		Do While Counts < 20 ' Bucle que acaba si Counts>=20 o con Exit Do.
			Counts = Counts + 1 ' Se incrementa Counts.
			If Counts = 10 Then ' Si Counts es 10.
				Check = False ' Se asigna a Check el valor False.
				Exit Do ' Se acaba el segundo Do.
			End If
		Loop
	Loop Until Check = False ' Salir del "loop" si Check es False.

	
	'Bucle WHILEWEND	
	Counts = 0
	While Counts < 20 
		Counts = Counts + 1 
	Wend
	
	'Bucle FOR EACH
	
	'found = false
	
	'For Each thisObject As String In thisCollection
	 '   If thisObject = "Hello" Then
	  '      found = True
	   '     Exit For
	   ' End If
	'Next thisObject
	
	'counter = 0
	
	'While counter < 20
    '	counter = counter + 1
    	' Insert code to use current value of counter.
	'End While
	
	
	'While counter <= 10
	'	Debug.Write(counter.ToString & " ")
	'	counter = counter + 1
	'End While
	
End Sub 
 */
	}

	private void Shell ( ) {

/* 
 
Private Sub m_WaitForm_Timer()
    Dim iExit As Long
    Dim hProg As Long
    
        
    Dim aux As Long
    'Dim number As Integer = 8
    'Dim index As Integer = 0
    Dim Counts As Integer
    Dim found As Boolean 
	Dim thisCollection As New Collection
    Dim counter As Integer
    
    
    hProg = OpenProcess(PROCESS_ALL_ACCESS, False, m_ProgHandle)
    
    GetExitCodeProcess hProg, iExit
    
    CloseHandle hProg
    
    
    'Sentencia IF
    If iExit <> STILL_ACTIVE Then
        m_ExitCode = iExit
        m_WaitForm.Visible = False
        Set m_WaitForm = Nothing
    End If
    
    
    'Bucle FOR
	FOR i=0 TO 6 STEP 2
		aux = i
	NEXT 

	
	'Sentencia CASE
	Select Case number
	Case 1 To 5
	    Debug.WriteLine("Between 1 and 5, inclusive")
	    ' The following is the only Case clause that evaluates to True.
	Case 6, 7, 8
	    Debug.WriteLine("Between 6 and 8, inclusive")
	Case IS = 9, IS = 10
	    Debug.WriteLine("Equal to 9 or 10")
	Case Else
	    Debug.WriteLine("Not between 1 and 10, inclusive")
	End Select

	'Sentencias DO LOOP UNTIL y DO WHILE LOOP
	Check = True ' Se inicializan las variables.
	Counts = 0
	Do ' Empieza sin comprobar ninguna condición.
		Do While Counts < 20 ' Bucle que acaba si Counts>=20 o con Exit Do.
			Counts = Counts + 1 ' Se incrementa Counts.
			If Counts = 10 Then ' Si Counts es 10.
				Check = False ' Se asigna a Check el valor False.
				Exit Do ' Se acaba el segundo Do.
			End If
		Loop
	Loop Until Check = False ' Salir del "loop" si Check es False.

	
	'Bucle WHILEWEND	
	Counts = 0
	While Counts < 20 
		Counts = Counts + 1 
	Wend
	
	'Bucle FOR EACH
	
	'found = false
	
	'For Each thisObject As String In thisCollection
	 '   If thisObject = "Hello" Then
	  '      found = True
	   '     Exit For
	   ' End If
	'Next thisObject
	
	'counter = 0
	
	'While counter < 20
    '	counter = counter + 1
    	' Insert code to use current value of counter.
	'End While
	
	
	'While counter <= 10
	'	Debug.Write(counter.ToString & " ")
	'	counter = counter + 1
	'End While
	
End Sub 
 */
	}

	private void Counts ( ) {

/* 
 
Private Sub m_WaitForm_Timer()
    Dim iExit As Long
    Dim hProg As Long
    
        
    Dim aux As Long
    'Dim number As Integer = 8
    'Dim index As Integer = 0
    Dim Counts As Integer
    Dim found As Boolean 
	Dim thisCollection As New Collection
    Dim counter As Integer
    
    
    hProg = OpenProcess(PROCESS_ALL_ACCESS, False, m_ProgHandle)
    
    GetExitCodeProcess hProg, iExit
    
    CloseHandle hProg
    
    
    'Sentencia IF
    If iExit <> STILL_ACTIVE Then
        m_ExitCode = iExit
        m_WaitForm.Visible = False
        Set m_WaitForm = Nothing
    End If
    
    
    'Bucle FOR
	FOR i=0 TO 6 STEP 2
		aux = i
	NEXT 

	
	'Sentencia CASE
	Select Case number
	Case 1 To 5
	    Debug.WriteLine("Between 1 and 5, inclusive")
	    ' The following is the only Case clause that evaluates to True.
	Case 6, 7, 8
	    Debug.WriteLine("Between 6 and 8, inclusive")
	Case IS = 9, IS = 10
	    Debug.WriteLine("Equal to 9 or 10")
	Case Else
	    Debug.WriteLine("Not between 1 and 10, inclusive")
	End Select

	'Sentencias DO LOOP UNTIL y DO WHILE LOOP
	Check = True ' Se inicializan las variables.
	Counts = 0
	Do ' Empieza sin comprobar ninguna condición.
		Do While Counts < 20 ' Bucle que acaba si Counts>=20 o con Exit Do.
			Counts = Counts + 1 ' Se incrementa Counts.
			If Counts = 10 Then ' Si Counts es 10.
				Check = False ' Se asigna a Check el valor False.
				Exit Do ' Se acaba el segundo Do.
			End If
		Loop
	Loop Until Check = False ' Salir del "loop" si Check es False.

	
	'Bucle WHILEWEND	
	Counts = 0
	While Counts < 20 
		Counts = Counts + 1 
	Wend
	
	'Bucle FOR EACH
	
	'found = false
	
	'For Each thisObject As String In thisCollection
	 '   If thisObject = "Hello" Then
	  '      found = True
	   '     Exit For
	   ' End If
	'Next thisObject
	
	'counter = 0
	
	'While counter < 20
    '	counter = counter + 1
    	' Insert code to use current value of counter.
	'End While
	
	
	'While counter <= 10
	'	Debug.Write(counter.ToString & " ")
	'	counter = counter + 1
	'End While
	
End Sub 
 */
	}

}
