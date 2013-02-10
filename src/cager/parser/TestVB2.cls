public Type FieldInformation ' Used To Hold Information Extracted From Table ti_tbolpfld
    FieldId As String
    Title As String
    PrintInfo As String
    DefaultText As String
    ValidationRule As Integer
    enum as long
End Type


Public Sub X()
1000000   	MsgBox
End Sub

private Function Fn()
1000010   	MsgBox
End Function

public property get pg() as Integer
1000020   	MsgBox
End Property

' TODO - this will fail in the VBErrHandlerVistor.
private sub xxx
End Sub

Private Sub CheckIf()
1000030   	dim x$
1000040   	dim y#
1000050   	dim z!
1000060   	set property = x
1000070   	Clause1
1000080   	select case CaseExpr
1000090   		case c1
1000100   			Clause1
1000110   		case is > 22
1000120   			Clause2
1000130   		case > 23
1000140   		case c3 TO C4, c5
1000150   			Clause345
1000160   		case else
1000170   			ClauseElse
1000180   	end Select
1000190       MsgBox "Hello":
    
'    If x then MsgBox "Hello":
1000200       Select Case X
1000210       Dim yy: Dim zz:
1000220       Case 0: MsgBox "0"
1000230       Dim fff
1000240       Dim yqfa: Case 1: MsgBox 1
1000250       End Select
End Sub


Private sub CheckPrimaryExpr()
1000260       xxLogError "(InsertFax) Can’t open or write to result set - Fax not sent"
1000270       Meth
1000280       Meth a, b, , d
1000290       Meth  , b, , d
1000300       Meth a, .b, c
1000310       Meth .a, b, c
1000320       Meth.SubMeth .a * b
1000330       Me.Print "Hello"
1000340       Print "Again"
1000350       xxx = rs!yyy
1000360       x = rs!!yy
1000370       With new zzz
1000380       End With
1000390       x = rsp![rate 1]
1000400       rsp![rate 1] = 44
1000410       Thing = rs![A funny thing]
1000420       line a b c !!
1000430       x = Trim(Property)
1000440       x = 5656@
1000450       call root.Interface(1)
1000460       x = 0.1!
1000470       x = Input(4)
1000480       x = Input(1, #1)
1000490       Input #1, 4
End Sub

private sub M(byval Event as long)
End Sub

private sub MM(Byval Property as integer)
End Sub

Private Sub Proc1()
1000500       ReDim LBan&(8), xx(3)
1000510       Label1:  Call pr_headings                        ' BUS
1000520       x = 22&
1000530       Label2:
1000540           event_cis_id = gl_cis_id                        ' BUS
1000550           If TrUe Then
1000560           	MsgBox 55
1000570           End If
        
        1000580  Label3:  If x Then
        1000590 	MsgBox 
        1000600    Else
        1000610 	MsgBox3
        1000620    End If
1000630      do
1000640      	MsgBox 11
1000650      loop
end sub


public function X() as String()
1000660   	msgbox
end function


Private Sub PrintAt(ByVal Text As String, Optional ByVal Col As clsReportColumn, Optional ByVal Y As Single = -1)
1000670       Dim X As Single
1000680       If Not Col Is Nothing Then
1000690           X = Col.ColumnLeft + Col.LeftMargin
1000700       Else
1000710           X = Printer.CurrentX
1000720       End If
1000730       With Printer
1000740           If X >= 0 Then .CurrentX = X
1000750           If Y >= 0 Then .CurrentY = Y
1000760           If X + Printer.TextWidth(Text) + 1 > PageWidth Then
1000770               X = PageWidth - Printer.TextWidth(Text) - 1
1000780           End If
1000790           If Y + Printer.TextHeight(Text) + 1 > PageHeight Then
1000800               Y = PageHeight - Printer.TextHeight(Text) - 1
1000810           End If
1000820           Printer.Print Text
1000830       End With
End Sub

Private Sub RightPrintAt(ByVal Text As String, Optional ByVal Col As clsReportColumn, Optional ByVal Y As Single = -1)
    ' Print Right-Aligned
1000840       Dim X As Single
1000850       If Not Col Is Nothing Then
1000860           X = Col.ColumnLeft + Col.ColumnWidth - Col.RightMargin - Printer.TextWidth(Text)
1000870       Else
1000880           X = Printer.CurrentX
1000890           If X + Printer.TextWidth(Text) + 1 > PageWidth Then
1000900               X = PageWidth - Printer.TextWidth(Text) - 1
1000910           End If
1000920       End If
1000930       With Printer
1000940           If X >= 0 Then .CurrentX = X
1000950           If Y >= 0 Then .CurrentY = Y
1000960           If Y + Printer.TextHeight(Text) + 1 > PageHeight Then
1000970               Y = PageHeight - Printer.TextHeight(Text) - 1
1000980           End If
1000990           Printer.Print Text
1001000       End With
End Sub

Private Function QueryDb() As RecordSet
1001010       Dim rs As RecordSet
1001020       On Error GoTo Handler

'    Set rs = dbAcc.OpenRecordset( _
'        "TRANSFORM Sum(PurchaseE.Net) AS Net " & _
'        "SELECT PurchaseE.PId, PurchaseE.Quarter, PurchaseE.Dte, PurchaseE.Supplier, PurchaseE.Cheque, PurchaseE.InvNum, PurchaseE.Bank, PurchaseE.Cash, PurchaseE.VAT, PurchaseE.Comment " & _
'        "From PurchaseE " & _
'        "GROUP BY PurchaseE.PId, PurchaseE.Quarter, PurchaseE.Dte, PurchaseE.Supplier, PurchaseE.Cheque, PurchaseE.InvNum, PurchaseE.Bank, PurchaseE.Cash, PurchaseE.VAT, PurchaseE.Comment " & _
'        "ORDER BY PurchaseE.CatName " & _
'        "PIVOT PurchaseE.CatName In ( " & _
'        "Pension, " & _
'        "[Postage and Stationary], " & _
'        "Salary,  " & _
'        "HMCE, " & _
'        "HMCT, " & _
'        "[Travel + Accom], " & _
'        "Sundries, " & _
'        "Accounts, " & _
'        "Payroll, " & _
'        "[Director Loan Account],  " & _
'        "Assets)")
1001030       Set rs = dbAcc.OpenRecordset(SQL, dbOpenSnapshot)
1001040       Set QueryDb = rs
1001050       Exit Function
1001060   Handler:
1001070       MsgBox "Could not read database: " & Err.Description, vbOKOnly + vbCritical
End Function

Private Function GetCategories() As Collection
1001080       Dim rs As RecordSet
1001090       Dim Coll As New Collection
1001100       Dim Cat As clsCategory
1001110       Set rs = dbAcc.OpenRecordset("Select * from Category Order by ReportPriority", , dbOpenSnapshot)
1001120       With rs
1001130           Do While Not .EOF
1001140               Set Cat = New clsCategory
1001150               Cat.CatId = !CatId
1001160               Cat.Catname = !Catname
1001170               Cat.PayRollInd = !PayRollInd
1001180               Cat.ReportHeading = !ReportHeading
1001190               Cat.ReportPriority = !ReportPriority
1001200               Coll.Add Cat
1001210               .MoveNext
1001220           Loop
1001230       End With
1001240       Set GetCategories = Coll
End Function

Private Sub CreateCols()
1001250       Dim Col As clsReportColumn
1001260       Dim Cat As clsCategory
    
    ' TODO
1001270       Const IncludeUnused = True
1001280       Set Columns = New clsReportColumns
1001290       With Columns
1001300           Set Col = .AddCol(0, DATE_WIDTH, PageWidth - RIGHT_MARGIN, LEFT_MARGIN, "Date", , , vbCenter, "dd/mm/yyyy")
1001310           Set Col = .AddCol(0, SUPPLIER_WIDTH, PageWidth - RIGHT_MARGIN, LEFT_MARGIN, "Supplier", , , vbLeftJustify, "")
1001320           Set Col = .AddCol(0, NOTE_WIDTH, PageWidth - RIGHT_MARGIN, LEFT_MARGIN, "", , , vbRightJustify, "(0)")  ' NOTE
1001330           Set Col = .AddCol(0, CHQ_WIDTH, PageWidth - RIGHT_MARGIN, LEFT_MARGIN, "CHQ", , , vbRightJustify, "")
1001340           Set Col = .AddCol(0, INV_WIDTH, PageWidth - RIGHT_MARGIN, LEFT_MARGIN, "INV", , , vbRightJustify)
1001350           Set Col = .AddCol(0, BANK_WIDTH, PageWidth - RIGHT_MARGIN, LEFT_MARGIN, "Bank", , , vbRightJustify, "£0.00")
1001360           Set Col = .AddCol(0, CASH_WIDTH, PageWidth - RIGHT_MARGIN, LEFT_MARGIN, "Cash", , , vbRightJustify)
1001370           Set Col = .AddCol(0, VAT_WIDTH, PageWidth - RIGHT_MARGIN, LEFT_MARGIN, "VAT", , , vbRightJustify)
1001380           For Each Cat In Categories
            ' TODO - if IncludeUnused=false, check one exists
            'PurchRs.FindFirst
1001390               Set Col = .AddCol(Cat.CatId, COL_WIDTH, PageWidth - RIGHT_MARGIN, LEFT_MARGIN, Cat.ReportHeading, , , vbRightJustify)
1001400           Next
1001410           Columns("Supplier").RightBorder = False
1001420       End With
1001430       For Each Col In Columns
1001440           Debug.Print "Col " & Col.Heading & " " & Col.ColumnLeft & "-" & Col.ColumnLeft + Col.ColumnWidth & " " & Col.Page
1001450       Next
End Sub


Private Sub PrintPages()
1001460       Dim Row As Integer
1001470       Dim RowY As Single
1001480       Dim RightEdge As Single
1001490       Dim i As Integer
1001500       Dim Col As clsReportColumn
1001510       Dim FontSize As Single
1001520       Dim MaxHPage As Integer
1001530       Dim NewPageRequired As Boolean
1001540       MaxHPage = Columns.Item(Columns.Count).Page
1001550       PageRow = 0
    
    ' Print Down and then Across:
1001560       For HPage = 1 To MaxHPage
1001570           With PurchRs
1001580               VPage = 0
1001590               PurchRs.MoveFirst
1001600               Row = 0
            
            ' Calculate right edge of this page
1001610               For Each Col In Columns
1001620                   If Col.Page = HPage Then
1001630                       RightEdge = Col.ColumnLeft + Col.ColumnWidth
1001640                   End If
1001650               Next
1001660               NewPageRequired = True
1001670               Do
1001680                   Row = Row + 1
                
                ' Check for end of page
1001690                   If .EOF Or NewPageRequired Then
1001700                       If VPage > 0 Then
1001710                           EndOfRows Columns, HPage, RowY, RightEdge
1001720                       End If
1001730                       If .EOF Then Exit Do
1001740                       VPage = VPage + 1
1001750                       StartOfPage HPage, VPage, Columns
1001760                       Row = 1
1001770                       Set NoteColl = New Collection
1001780                       NewPageRequired = False
1001790                   End If
1001800                   RowY = (Row - 1) * LineHeight + TABLE_BODY_TOP
1001810                   If HPage = 1 Then
1001820                       Columns(1).PrintValue !Dte, RowY
1001830                       Columns(2).PrintValue !Supplier, RowY
1001840                       If !Comment <> "" Then
1001850                           NoteColl.Add !Comment
1001860                           Columns(3).PrintValue CStr(NoteColl.Count), RowY
1001870                       End If
1001880                       If !Cheque <> "" Then Columns(4).PrintValue !Cheque, RowY
1001890                       If !InvNum <> "" Then Columns(5).PrintValue !InvNum, RowY
1001900                       If !Bank <> 0 Then Columns(6).PrintValue !Bank, RowY, , True
1001910                       If !Cash <> 0 Then Columns(7).PrintValue !Cash, RowY, , True
1001920                       If !VAT <> 0 Then Columns(8).PrintValue !VAT, RowY, , True
1001930                   End If
                
                ' Check if CatId is in the current column range
1001940                   For i = 1 To Columns.Count
1001950                       If Columns(i).DbId = !CatId Then
1001960                           If Columns(i).Page = HPage Then
1001970                               Columns(i).PrintValue !Net, RowY, , True
1001980                           End If
1001990                           Exit For
1002000                       End If
1002010                   Next
1002020                   Printer.DrawWidth = 1
1002030                   Printer.DrawMode = vbBlackness
1002040                   Printer.DrawStyle = vbSolid
1002050                   RowY = RowY + Printer.TextHeight("M") + (1 * PixelY)
                ' TODO (?) - go all the way to the right for page 1, all the way to the left for page 2. . .
                'if if x then exit sub
1002060                   If Row > 40 Then
                    ' TEMP
1002070                       .MoveLast
                    'NewPageRequired = True
1002080                   End If
1002090                   MoveNext
1002100               Loop
1002110           End With
1002120       Next
End Sub

Private Sub StartOfPage( _
            ByVal HPage As Integer, _
            ByVal VPage As Integer, _
            ByVal Columns As clsReportColumns)
1002130       Dim Col As clsReportColumn
1002140       Dim FontSize As Single
1002150       If HPage <> 1 Or VPage <> 1 Then Printer.NewPage
1002160       For Each Col In Columns
1002170           If Col.Page = HPage Then
1002180               FontSize = Printer.Font.Size
1002190               Printer.Font.Size = Col.HeaderFontSize
1002200               Col.PrintValue Col.Heading, TABLE_HEADING_TOP + 0.2
1002210               Printer.Font.Size = FontSize
1002220           End If
1002230       Next
    
End Sub
            

Private Sub EndOfRows( _
            ByVal Columns As clsReportColumns, _
            ByVal HPage As Long, _
            ByVal LastRowY As Single, _
            ByVal RightEdge As Single)
1002240       Dim X1 As Single, Y1 As Single
1002250       Dim X2 As Single, Y2 As Single
1002260       Dim OldWidth As Long
1002270       Dim Col As clsReportColumn
1002280       OldWidth = Printer.DrawWidth
1002290       X1 = LEFT_MARGIN
1002300       X2 = RightEdge
1002310       Y1 = TABLE_HEADING_TOP - (6 * PixelY)
1002320       Y2 = LastRowY + 0.05
1002330       Printer.DrawWidth = 3
    ' Draw line above subtotals
    'Printer.Line (X1, Y2)-(X2, Y2)
1002340       Y2 = Y2 + SUBTOTAL_GAP_ABOVE
1002350       Columns.PrintSubtotals Y2, HPage
    
    ' Draw line beneath subtotals
1002360       Y2 = Y2 + Printer.TextHeight("M") + SUBTOTAL_GAP_BELOW
    'Printer.Line (X1, Y2)-(X2, Y2)
1002370       Printer.DrawWidth = 1
    ' Draw the right-boundary of each column (including subtotals)
1002380       For Each Col In Columns
1002390           X1 = Col.ColumnLeft + Col.ColumnWidth '+ col.LeftMargin
1002400           If HPage = Col.Page And Col.RightBorder Then
            'Printer.Line (X1, Y1)-(X1, Y2)
1002410           End If
1002420       Next
    
    ' Draw Box around whole thing
1002430       X1 = LEFT_MARGIN
1002440       Printer.DrawWidth = 5 ' pixels
    'Printer.Line (X1, Y1)-(X2, Y2), , B
    
    ' Draw a line under headings
1002450       Printer.DrawWidth = 5 ' pixels
1002460       Y1 = TABLE_BODY_TOP - (20 * PixelY)
    'Printer.Line (X1, Y1)-(X2, Y1)
1002470       Printer.DrawWidth = OldWidth
End Sub


Public Sub PrintPurchases()
1002480       Set Categories = GetCategories()
1002490       Set PurchRs = QueryDb()
1002500       If PurchRs Is Nothing Then Exit Sub
1002510       With Printer
1002520           .ScaleMode = vbCentimeters
1002530           .Font.Name = "Times New Roman"
1002540           .Font.Size = 8
1002550           .DrawWidth = 2 ' Pixels
1002560           PixelX = .ScaleX(1, vbPixels)
1002570           PixelY = .ScaleY(1, vbPixels)
        
        ' LineHeight = TextHeight*1.2 + 3 pixels (allow for horiz line).
1002580           LineHeight = .TextHeight("A") * 1.2 + (3 * PixelY)
1002590           PageHeight = Printer.ScaleHeight
1002600           PageWidth = Printer.ScaleWidth
1002610       End With
1002620   
SomeOtherLabel:
1002630       CreateCols
1002640       PrintPages
1002650       Printer.EndDoc
End Sub


Private Sub Class_Initialize()
1002660       SQL = "Select PId, Quarter, Dte, Cheque, InvNum, Bank, Cash, VAT, Supplier, CatName, CatId, Net, Comment from PurchaseE"
End Sub

