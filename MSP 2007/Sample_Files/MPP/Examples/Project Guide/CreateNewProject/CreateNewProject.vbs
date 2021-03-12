'***********************************************************************
' NAME: init
' PURPOSE: Here we set up the sidepane.
' This gets called when the sidepane first loads.
' PARAMETERS: None
' RETURN: None
'***********************************************************************
sub init
   'Suppress any excess script errors to avoid user confusion.
    on error resume next
    pSetupSidepane(true)
end sub

'***********************************************************************
' NAME: validateData
' PURPOSE: This will validate the current values of:
'		Project Title
'		Comments
'		Selected Template
' PARAMETERS: None
' RETURN: None
'***********************************************************************
sub validateData()
	dim currentApplication
	dim templateName
	dim errorMessage
	
   'Suppress any excess script errors to avoid user confusion.
    on error resume next
    
   'Check for blank data 
    errorMessage = ""
	if trim(ProjectTitle.value)="" then 
		errorMessage = "   Enter a Project Title." & chr(13) & chr(10)
		ProjectTitle.value = ""
	end if
	if trim(ProjectNotes.value)="" then 
		errorMessage = errorMessage & "   Enter a Project Description." & chr(13) & chr(10)
		ProjectNotes.value=""
	end if
	

    'Set template name based on the Project Type option selected 	
	select case true
	case SelectTemplate(0).checked 
		templateName = "C:\Program Files\Microsoft Office\Templates\1033\PROJOFF.MPT"
	case SelectTemplate(1).checked 
		templateName = "C:\Program Files\Microsoft Office\Templates\1033\NEWBIZ.MPT"
	case SelectTemplate(2).checked 
		templateName = "C:\Program Files\Microsoft Office\Templates\1033\INFSTDEP.MPT"
	case SelectTemplate(3).checked
		templateName = "C:\Program Files\Microsoft Office\Templates\1033\SOFTDEV.MPT"
	case else
		errorMessage = errorMessage & "   Select a Project Type." & chr(13) & chr(10)
	end select


	if errorMessage <> "" then 
		errorMessage = "To continue you must:" & chr(13) & chr(10) & chr(13) & chr(10) & errorMessage & chr(13) & chr(10)
		msgbox  errorMessage,,"Project Guide"
		exit sub
	end if
	set currentApplication = window.external.application
	currentApplication.FileOpen templateName,,,,,,,,,"MSProject.MPT"
	currentApplication.Activeproject.Title = ProjectTitle.value
	currentApplication.Activeproject.ProjectSummaryTask.Notes = ProjectNotes.value
	pNavigate 1,-1,"createPlan"
end sub
