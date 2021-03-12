Attribute VB_Name = "Module1"
Sub FilterCriticalTasksByDateAndResource()
Dim strResourceName As String

'Macro FilterCriticalTasksByDateAndResource
'Macro Recorded Sun 5/4/03 by Administrator.
    
    strResourceName = InputBox("Please enter the resource name:")
    FilterEdit Name:="Late Tasks On Critical Path", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Finish", Test:="is less than", Value:="Mon 9/15/03", ShowInMenu:=False, ShowSummaryTasks:=True
    FilterEdit Name:="Late Tasks On Critical Path", TaskFilter:=True, FieldName:="", NewFieldName:="Critical", Test:="equals", Value:="Yes", Operation:="And", ShowSummaryTasks:=True
    FilterEdit Name:="Late Tasks On Critical Path", TaskFilter:=True, FieldName:="", NewFieldName:="Resource Names", Test:="contains", Value:=strResourceName, Operation:="And", ShowSummaryTasks:=True
    FilterApply Name:="Late Tasks On Critical Path"
End Sub

