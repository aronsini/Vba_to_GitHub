# About Vba_to_GitHub
[![Paypal donation](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.com/donate?hosted_button_id=8KKENL9VDW9GL)

Module to manage modules between GitHub and VBA. Composed of basic processes but of significant usefulness for developers.

Through this module you will be able to upload, download, update and delete files from a repository on GitHub, as well as processes for manipulating code from VBE.

# References

- Microsoft Visual Basic for Applications Extensibility 5.3
- Microsoft XML, v6.0
- Microsoft Office 16.0 Object Library

# Put VBA module from GitHub file

Once Vba_to_GitHub are in your project, you could add any file content on GitHub as module code on VBA, uou could take any file on GitHub and the macro do the rest, the result is the content of link as string so if the code is not writed on VBA then you could have errors. Example below.

```VB.net
Sub CreateModuleFromGitHubFileContent()
 
    Dim strFileUrl as String
    strFileUrl = https://github.com/VBA-tools/VBA-Web/blob/master/README.md
   
    Dim wbMyProject as Workbook
    Set wbMyProject = Application.ActiveWorkbook
    
    'If any workbook is defined or passed as argument to Vba_AddGitModule proc, then the module will be added to personal macros workbook.
    Vba_AddGitModule strFileUrl, wbMyProject
    
End function
```

# Update

If you want to update this module directly from the code you only need to run the next proc on any module with your own project workbook.

```VB.net
Function UpdateMyGitModule(Origin As String, Destination As String) As String
 
End function
```

# Donation
Open source development, but any help is welcome. Donations are in USD so take it in mind if not is your local currency.

[![Paypal donation](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/donate?hosted_button_id=8KKENL9VDW9GL)
