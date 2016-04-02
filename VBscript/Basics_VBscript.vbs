'==========================================================================
' SYNOPSIS
' 	Wscript script template
'
' DESCRIPTION
' 	
'
' NOTES 
'    File Name  : xxxxx.vbs
'    Author     : Janice McCullough
'    Requires   : Wscript 5.6
'
' USEFUL LINKS:
'
'			
'==========================================================================
option explicit
dim name, veggie, action, character

'declare variables
name = "Peter"
veggie = " Pumpkin"
action = " Eater"

character = name & veggie & action
wscript.echo character


wscript.echo "Done"