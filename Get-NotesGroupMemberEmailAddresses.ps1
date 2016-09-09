Function Get-NotesGroupMemberEmailAddresses {
    
<#
.Synopsis
Translates Lotus Notes Group Members into Internet Address/Primary SMTP Address


.Description
Queries Lotus Notes address book (names.nsf) to get SMTP email address and object type (User, group, mailin db, etc) for all group members.

NOTE - this function must be run from a 32 bit (x86) Powershell session if your Lotus Notes client is 32 bit.


.PARAMETER NotesName
Mandatory. This is the name of the Group within Lotus Notes you want to get information on. Use just the name without any prefix or suffix.

Valid: "Test Group"

Invalid: "Test Group/CompanyName"
    
        
.Example
$NotesGroup = "Test Group" ; $NotesGroup | Get-NotesGroupMemberEmailAddresses

Sample of calling the function using pipeline input
     
        
.Example
Get-NotesGroupMemberEmailAddresses "Test Group"

Sample of calling the function using the positional parameter "NotesName" - the value supplied is assumed to be intended for this parameter
      

.Example     
Get-NotesGroupMemberEmailAddresses -NotesName "Test Group"

Sample of calling the function and explicitly specifying the parameter "NotesName" and the value for it


.Example
(Get-NotesGroupMemberEmailAddresses "Test Group").EmailAddress

Sample of how to get just the EmailAddress result from the function
            

.Example
Get-NotesGroupMemberEmailAddresses "Test Group" | select EmailAddress

Another way to get just the EmailAddress result from the function
                
#>

    
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline=$True)]
        [string[]]$NotesName
        )

    Begin {
        
        $notes = new-object -comobject Lotus.NotesSession
        $notes.initialize('') 
        $dbNAB = $notes.GetDatabase( 'YourServerNameGoesHere', 'names.nsf', 1 ) 
        $vwUsers = $dbNAB.GetView( '($Users)' ) 
        
        $GroupMemberResults = @()
        
        } #end of begin block

    Process {
        
        Foreach ($name in $NotesName) {

            Try {
                
                $doc = $vwUsers.getdocumentbykey($name,$true)
                
                $GroupMembers = $doc.ColumnValues[4]

                ForEach ($GroupMember in $GroupMembers) {
                    
                    $GroupMember = ($GroupMember -split "@")[0]
                                                           
                    $GroupMemberDoc = $vwUsers.getdocumentbykey($GroupMember,$true)
                
                    $SmtpAddress = ($GroupMemberDoc.ColumnValues[16]).Trim()

                    $ObjectType = ($GroupMemberDoc.ColumnValues[25]).Trim()
                    
                    $TempObject = $null

                    $TempObject = New-Object PSObject

                    Add-Member -InputObject $TempObject -MemberType NoteProperty -Name "EmailAddress" -Value $SmtpAddress

                    Add-Member -InputObject $TempObject -MemberType NoteProperty -Name "MemberType" -Value $ObjectType 

                    $GroupMemberResults += $TempObject
                                                            
                    } # end of ForEach loop        
                
                } # end of try block
        
            Catch{

                Write-Error $_.Exception

                } # end of catch block

        } # end of ForEach

    } # end of process block

    End {
        
        $notes = $null

        Write-Output $GroupMemberResults
        
        } # end of end block

    } # end of Get-NotesGroupMemberEmailAddresses function