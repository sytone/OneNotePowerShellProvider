##
##  Enumerate existing OneNote notebooks
##    
@{
     Name='Enumeration';
     Description=$(@'
You can enumerate the current OneNote notebooks using the OneNote: PSDrive.
'@);
     ScriptBlock={
            dir OneNote:\
     };
}
    
@{
     Name='Create notebook structure';
     Description=$(@'
You can create notebooks, sections, and pages with the New-Item cmdlet.
Right now, I don't support section groups.
'@);
     ScriptBlock={


            new-item -path onenote:\TempNotebook -type Notebook -value "$(gc env:\temp)"
            new-item -path OneNote:\TempNotebook\Section -type section
            new-item -path "OneNote:\TempNotebook\Section\Page 1" -type page
            new-item -path "OneNote:\TempNotebook\Section\Page 2" -type page
            new-item -path "OneNote:\TempNotebook\Section\Page 3" -type page
            new-item -path "OneNote:\TempNotebook\Section 2" -type section

            ##
            ##  Note that relative names work. Also note that you can abbreviate the type name.
            ##

            pushd "OneNote:\TempNotebook\Section 2"
            new-item -path ".\Page 1" -type p
            new-item -path ".\Foobar" -type p
            new-item -path ".\Razzle" -type p
            new-item -path OneNote:\TempNotebook\Foo -type section
            dir OneNote:\TempNotebook -recurse
            popd
        

     };
     
}
    
@{
     Name='Names with slashes...';
     Description=$(@'
Because forward slashes and back slashes are used as path separators, and ':' is the drive separator, 
        you need to escape them when using them in the names of pages, sections, etc.
'@);
     ScriptBlock={


            ##
            ##  Note that OneNote paths are not case-sensitive.
            ##

            new-item -path "onenote:\tempnotebook\section\either&amp;#47;or" -type page
            new-item -path "onenote:\tempNotebook\section\C&amp;#58;&amp;#92;Users&amp;#92;BDewey" -type page
            dir onenote:\tempnotebook\section 
        

     };
     
}
    
@{
     Name='Putting text on pages';
     Description=$(@'
You can use the add-content cmdlet to put text on pages. You can use "type" or "get-content"
            to see the OneNote XML that makes up the page.
'@);
     ScriptBlock={


            get-process | out-string | add-content "onenote:\TempNotebook\Section\Page 1"
            type "onenote:\TempNotebook\Section\Page 1"
        

     };
     
}
    
@{
     Name='Putting attachments on pages';
     Description=$(@'
If you pipe a FileInfo object to add-content, then the file will be embedded on the page.
'@);
     ScriptBlock={


            dir "$env:SystemRoot\System32\oobe\en-US\privacy.rtf" | add-content "OneNote:\TempNotebook\Section\Page 2"
            type "OneNote:\TempNotebook\Section\Page 2"
        

     };
     
}
    
@{
     Name='Replacing page content';
     Description=$(@'
You can use set-content to replace the content of a page.
'@);
     ScriptBlock={


            "No more process listing!" | set-content "onenote:\TempNotebook\Section\Page 1"
            type "onenote:\TempNotebook\Section\Page 1"
        

     };
     
}
    
@{
     Name='Getting the OneNote hierarchy';
     Description=$(@'
You can us the Get-OneNoteHierarchy cmdlet to get the XML hierarchy.
'@);
     ScriptBlock={


            get-command Get-OneNoteHierarchy | select-object Definition | format-list

            ##
            ##  This will return the hierarchy as a parsed [xml] object.
            ##

            Get-OneNoteHierarchy OneNote:\TempNotebook

            ##
            ##  ...whereas this will return the XML text. Note the use of "scope" to determine
            ##  how much to show.
            ##

            Get-OneNoteHierarchy OneNote:\TempNotebook -scope hsSections -xml
        

     };
     
}
    
@{
     Name='Getting Hyperlinks';
     Description=$(@'
You can use the Get-OneNoteHyperlink cmdlet to get the hyperlinks to any item.
'@);
     ScriptBlock={


            Get-OneNoteHyperlink 'OneNote:\TempNotebook\Section\Page 2'

            ##
            ##  In addition to giving a OneNote path on the command line, you can pipe OneNote items
            ##  to the cmdlet.
            ## 

            dir OneNote:\TempNotebook | Get-OneNoteHyperlink
        

     };
     
}
    
@{
     Name='Exporting pages';
     Description=$(@'
You can use the Export-OneNote cmdlet to export OneNote content.
'@);
     ScriptBlock={


            export-onenote 'OneNote:\TempNotebook\Section\Page 2' ${env:TEMP}\TestPage.mht -format mht
            dir "${env:TEMP}\TestPage.mht"
        

     };
     
}
    
@{
     Name='Closing notebooks';
     Description=$(@'
The clear-item cmdlet will close notebooks.
'@);
     ScriptBlock={

clear-item OneNote:\TempNotebook

     };
     
}
    
@{
     Name='Cleanup...';
     Description=$(@'
Remove the test notebook.
'@);
     ScriptBlock={


            remove-item "$(gc env:\temp)\TempNotebook" -recurse -force
            del "${env:TEMP}\TestPage.mht"
        

     };
     
     ErrorAction='continue';
     
}
    
