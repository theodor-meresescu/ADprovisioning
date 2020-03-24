Set-ExecutionPolicy -ExecutionPolicy RemoteSigned

class Stage {
    [string] $Name
    hidden [DateTime] $StartTime = [DateTime]::Now

    Stage($Name) {
        $this.Name = $Name
    }

    [TimeSpan] GetElapsed(){
        return [DateTime]::Now - $this.StartTime
    }

    [string] GetHeader() {
        return "~ [$this.Name] ~"
    }
}

class Create : Stage {

    Create () : base('Creating user') { }

    [System.Object] CreateUser($user) {

        $randomPassword = (-join ((65..90) + (97..122) + (34..56) | Get-Random -Count 12 | % {[char]$_}))
        
        $sAMAccountName = $user.sAMAccountName
        $name = $user.Name
        $mail = $user.Email
        $UPN = $user.UPN
        $givenName = $name.Split(" ")[0]
        $surName = $name.Split(" ")[-1]

        try {
            New-ADUser -Name "$sAMAccountName" -DisplayName "$name" -GivenName "$givenName" -Surname "$surName" -EmailAddress $mail -SamAccountName $sAMAccountName -UserPrincipalName $UPN -Path "OU=_RootOU,DC=blah,DC=org" -AccountPassword(ConvertTo-SecureString $randomPassword -asplaintext -force) -Description "Auto-provisioned user." -Enabled $true
             
            $user | Add-Member -NotePropertyMembers @{ProvisioningStatus="0"
                    ProvisioningMessage="Created"} -PassThru
        }
        catch {
            $ErrorMessage = $Error[0].Exception[0].Message
            $ErrorTargetObject = $Error[0].TargetObject
            $ErrorCode = $Error[0].Exception[0].ErrorCode

            $user | Add-Member -NotePropertyMembers @{ProvisioningStatus="$ErrorCode"
                    ProvisioningMessage="$ErrorMessage : $ErrorTargetObject"} -PassThru
        }
       
        return $user
    }

    [System.Object] Invoke([UserFactory]$Job) {
        
        $Job.LogHeader($this.GetHeader())

        $FactoryCreateUsers = New-Object System.Collections.Generic.List[System.Object]

        $Job.Users | ForEach-Object {

            $newUser = CreateUser($_)

            $FactoryCreateUsers.Add($newUser)

            $Job.LogEntry("[in {0:N2}s]" -f $this.GetElapsed().TotalSeconds)
        }

        $Job.Users = $FactoryCreateUsers

        return $Job
    }
}

class UpdateGroup : Stage {

    UpdateGroup () : base('Adding to group') { }

    [System.Object] Invoke([UserFactory]$Job) {

        $memberOf = ""

        $Job.LogHeader($this.GetHeader())
        
        $Job.Users | ForEach-Object { 
            Add-ADGroupMember -Identity $memberOf -Members $_.sAMAccountName
            $_ | Add-Member -NotePropertyMembers @{MemberOf="$memberOf"} -PassThru
            $Job.LogEntry("[in {0:N2}s]" -f $this.GetElapsed().TotalSeconds)
        }

        return $Job
    }
}

class SendEmail : Stage {

    SendEmail () : base('Sending email') { }

    [System.Object] Invoke([UserFactory]$Job) {

        $template = (Get-Content "")
        [string] $smtpserver = ""
        [string] $fromAddr = ""
        [string] $subject = ""
        [string] $maintainer = ""
        [string] $message = ""

        $html = null

        $template | ForEach-Object { 
            if($_.Trim() -eq "<p>$[CONTENT]</p>") {
              -join(($_).ToString().Replace("$[CONTENT]", $message),$html)
              } else { -join($_, $html) }
            }

        $Job.Users | ForEach-Object { 
            Send-MailMessage -From $fromAddr -To $_.Email -Cc $maintainer -Subject $subject -Body ($html | Out-String) -BodyAsHtml -SmtpServer $smtpserver -Encoding UTF8
            $_ | Add-Member -NotePropertyName EmailSent -NotePropertyValue Done
            $Job.LogEntry("[in {0:N2}s]" -f $this.GetElapsed().TotalSeconds)
        }

        return $Job
    }
}

class UpdateDb : Stage {
    UpdateDb () : base('Updating database') { }

    hidden [string] $connectionString = ""

    [System.Data.SqlClient.SqlConnection] $Connection = [System.Data.SqlClient.SqlConnection]::new($this.connectionString)

    [string] UpdateUserStatus($ProvisionedUsers) {

        $sb = [Text.StringBuilder]::new()
        [void]$sb.AppendLine("MERGE INTO AD_Batch_Provisioning")
        [void]$sb.AppendLine("USING (")
        [void]$sb.Append("VALUES ")

        $ProvisionedUsers | ForEach-Object {
            [void]$sb.AppendLine("('" + $_["ProvisioningStatus"] + "', '" + $_["ProvisioningMessage"]  + "')")

            if($_ -ne $ProvisionedUsers[-1]) {
            [void]$sb.Append(",")
            }
        }

        [void]$sb.AppendLine("")
        [void]$sb.AppendLine(") AS source (ProvisioningStatus, ProvisioningMessage)")
        [void]$sb.AppendLine("ON AD_Batch_Provisioning.sAMAccountName = source.sAMAccountName")
        [void]$sb.AppendLine("WHEN MATCHED THEN")
        [void]$sb.AppendLine("UPDATE SET ProvisioningStatus = source.ProvisioningStatus, ProvisioningMessage = source.ProvisioningMessage;")
    

        return $sb.ToString()
    }

    [System.Object] Invoke([UserFactory]$Job) {

        [string] $userUpdateQuery = UpdateUserStatus($Job.Users)

        $this.Connection.Open()

        $command = $this.Connection.CreateCommand()
        $command.CommandText = $userUpdateQuery

        $command.ExecuteReader()

        $this.Connection.Close()

        $Job.LogEntry("[in {0:N2}s]" -f $this.GetElapsed().TotalSeconds)

        return $Job
    }
}

class UserFactory {

    [System.Collections.Generic.List[System.Object]] $Users

    hidden [array] $Result = @()
    hidden [DateTime] $StartTime = [DateTime]::Now
    hidden [System.Collections.Generic.List[Stage]] $Stages

    UserFactory ($toProvision) {
        $this.Users = $toProvision
    }

    [TimeSpan] GetElapsed(){
        return [DateTime]::Now - $this.StartTime
    }

    [void] LogHeader([string]$S) {
        $this.Result += $S
    }

    [void] LogEntry([string]$S) {
        $this.Result += "`t$S"
    }

    [void] LogError([string]$S) {
        $this.LogEntry("`!![$S]!!")
    }

    [UserFactory] AddStage([Stage]$S) {
        $this.Stages.Add($S)

        return $this
    }

    [UserFactory] Invoke() {
        $this.Stages | ForEach-Object {
            try {
                $this = $_.Invoke($this)
            }
            catch {
                $this.LogError($_.Exception.Message)
                break
            }
        }

        return $this
    }

    [string] GetResult() {
        return $this.Result | Out-String
    }
}

[UserFactory]::New($this.ToProvision)
.AddStage([Create]::New())
.AddStage([UpdateGroup]::New())
.AddStage([SendEmail]::New())
.AddStage([UpdateDb]::New()).Invoke().GetResult()
