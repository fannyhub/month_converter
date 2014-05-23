Sub month_counter()
Dim Datee, my_date, monthh, year_dic
Start:
Datee = InputBox("inserisci la data in formato dd/mm/yyyy", , Fix(Date))
If Datee = "" Then
Exit Sub
End If
On Error GoTo ErrHandler
my_date = CDate(Datee)
monthh = Month(my_date)

'Debug.Print monthh




Set year_dic = CreateObject("scripting.dictionary")
With year_dic
.Add 1, "gennaio"
.Add 2, "febbraio"
.Add 3, "marzo"
.Add 4, "aprile"
.Add 5, "maggio"
.Add 6, "giugno"
.Add 7, "luglio"
.Add 8, "agosto"
.Add 9, "settembre"
.Add 10, "ottobre"
.Add 11, "novembre"
.Add 12, "dicembre"

End With
Debug.Print year_dic(monthh)
MsgBox year_dic(monthh)

ErrHandler:
Resume Start:
End Sub
