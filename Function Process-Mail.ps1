<#
.Synopsis
	Process everything needed to maintain and send mail-related information. Support -whatif.
.Description
	Process everything needed to maintain and send mail-related information
	It uses a special [hashtable] variable in the global scope (default: $Mail)
	The body of mail messages is in html format. Some basic formatting styles are applied which can be overriden by corresponding css properties in the variable.
.Inputs
    System.String
.Outputs
    None
.Parameter Subject
	The subject line of the mail to be sent
.Parameter Reason
	The reason why this function is being called. Receivers can selectively subscribe to receive mails for certain reasons only.
	Don't specify 'List' as it is a reserved reason.
.Parameter Mail
	The mail-related information [hashtable] variable in the global scope, its default name is $Mail.

	$Mail.Sender			The string to be used in the mail "From" field.
							If omitted, it defaults to "$ScriptName <noreply.$Scriptname@$Env:ComputerName".
							The "From" field can be used at the receiver's input mailbox to assist in automatic mail filtering.

	$Mail.Receiver			The string array of receiver email-addresses, separated by comma or commapoint.

	$Mail.Server			The string array of mailserver names/ip-addresses:port(optional), separated by comma or commapoint.
							Maximum 3 attempts are made to contact every mailserver to deliver the message.

	$Mail.List				The string array of reasons for which a list must be maintained.
							The reason values are dynamically defined by calling this function with a '-Reason' parameter.

	$Mail.Style				Common css (inline) styles to be applied for different parts of the mail message.
							body:		Style(s) to be applied to the <body> tag
							h1:			Style(s) to be applied on the <h1> tag
							p:			Style(s) to be applied on the <p> tag
							table:		Style(s) to be applied on the <table> tag
							th:			Style(s) to be applied on the <th> tag
							td:			Style(s) to be applied on the <th> tag
							error:		Additional style(s) to be applied when a type [ErrorRecord] is passed as object or variable

							If any style is missing, a default style will be added to $Mail.Style and used for css formatting.
							Common styles may be overridden by receiver styles (personal preference or mail client support).
							
	$Mail.$Receiver.$Reason	The mail action (and its format) to be performed when this function is called with a -Reason parameter.
							If this hashtable key exists, its accepted values are 'Basic', 'Full', 'Text' and 'List'.
							Basic	The mail body contains the passed input objects/variables without type information.
							Full	The mail body contains the passed input objects/variables with type information.
									This format is best suited for receivers with a technical background.
									Distinction can be made between [string]'true' and [bool]true or [string]'1' and [int16]1.
							Text	The mail body contains the passed input without object/variable nor type information.
									This format is best suited for receivers with a nontechnical background.
							List	Indicates that this receiver receives the accumulated list for this $Reason.
							Formats may be abbreviated (such as Test=B,F,T,L or Test=BFTL => Test=Basic,Full,Text,List).

	$Mail.$Receiver.List	The format of the list being sent to this receiver.
							Basic	The mail body contains the list and any other input objects/variables without type information.
							Full	The mail body contains the list and any other input objects/variables with type information.
							Text	The mail body contains the list and any other input without object/variable nor type information.
							Formats may be abbreviated (such as List=B,F,T or List=BFT => List=Basic,Full,Text).

	$Mail.$Receiver.Style	Css (inline) styles to be applied for this receiver for different parts of the mail message.
							See $Mail.Style for the available options.

	$Mail.ListName			A string used as variable name to group the dynamically generated lists together.
							If omitted, it contains the timestamp (format HHmmssffff) when the function is called for the first time.
							Specifying the name of the mail variable (eg: Mail) is possible but can cause conflicts.

							Individual [ArrayList] lists are gathered based upon $Mail.List under the name $ListName.$Reason.
							Not every specified list may exist, as it is created only at the first occurrence of $Reason.

	$Mail.Signature			A (html-formatted) signature text to be used as final (or only) part of the mail body. 
							Its value is replaced every time -Signature is included in the function call.
							If $Mail.Signature exists, it is automatically included in the mail body.
							It can contain a PowerShell cmdlet to be executed (such as Get-Content).
							The signature text can contain dynamic elements (such as $(Get-Date) or $variables) substituted appropriately. 

	$Mail.SubjectPrefix		A string to be used to prefix the passed Subject parameter.
							Its value is replaced every time the -SubjectPrefix is included in the function call.
							Its presence in the function call automatically switches the -Prefix parameter to $true.
							The subject prefix can be used at the receiver's input mailbox to assist in automatic mail filtering.

	Example:	$Mail = @{}											# Define top level HashTable
				$Mail.Sender = '$MyScript <$MyScript@$MyHome>'		# Automatic variable substitution
				$Mail.Receiver = @('His.Name@HisMailbox.com','hername@hersite.org','Its.Name@itshost')
				$Mail.Server = @('LoadBalancer','MailServer1','MailServer2','103.6.0.5')
				$Mail.Signature = Get-Content -Path "C:\ProjectX\MailSignature.html"
				$Mail.SubjectPrefix = '$MyScript`: '				# Initial SubjectPrefix, inserted only when $Prefix:$true
				$Mail.List = @('test1','test2','test3')				# The reasons for which lists will be maintained
				$Mail.ListName = 'MailList'							# Variable name to group any mail lists
				$Mail.'His.Name@HisMailbox.com'=@{}					# Define nested HashTable 'His.Name@HisMailbox.com'
				$Mail.'His.Name@HisMailbox.com'.test1 = 'Full'		# His.Name@HisMailbox.com receives reason test1 in format Full
				$Mail.'His.name@HisMailbox.com'.test2 = 'B','List'	# His.Name@HisMailbox.com receives reason test2 in format Basic
				$Mail.'his.name@HisMailbox.com'.test4 = 'Basic'		# His.Name@HisMailbox.com recieves reason test4 in format Basic
				$Mail.'HIS.NAME@HISMAILBOX.COM'.list = 'F'			# His.Name@HisMailbox.com receives list test2 in format Full
				$Mail.'Hername@hersite.org'=@{}						# Define nested HashTable 'Hername@hersite.org'
				$Mail.'Hername@hersite.org'.test3 = 'list'			# hername@hersite.org receives list test3 in format Basic
				$Mail.'herName@hersite.org'.list = 'Basic'			# Defines the format of the lists being sent to this receiver
				$Mail.'Its.Name@itshost'=@{}						# Define nested HashTable 'Its.Name@itshost'
				$Mail.'Its.Name@itshost'.test1 = 'B','List'			# Missing $Mail.'Its.Name@itshost'.List format defaults to Text
				$Mail.'Its.Name@itshost'.test4 = 'List'				# Never sent because a list for reason test4 is not maintained (missing in $Mail.List)

				The real power of putting everything in one variable becomes obvious when the parameters are not hard-coded inside but kept outside the script. 
				External Powershell scripts (such as Get-IniFile) read the contents of an INI-file and produce a multilevel hashtable, capable of becoming the $Mail variable.

				Contents of 'C:\ProjectX\Mail.ini':
				Sender='$MyScript <noreply.$Myscript@$MyHome'
				Receiver=His.Name@HisMailbox.com,hername@hersite.org;Its.Name@itshost
				Server=LoadBalancer;MailServer1,MailServer2,103.6.0.5
				Signature=Get-Content -Path 'C:\ProjectX\MailSignature.html'
				SubjectPrefix='$MyScript`: '
				List=test1,test2,test3
				ListName=MailList
				[His.Name@HisMailbox.com]
				test1=Full
				test2=B,List
				test4=Basic
				list=F
				[hername@hersite.org]
				test3=list
				list=Basic
				[Its.Name@itshost]
				test1=B,List
				test4=List

				$Mail = Get-IniFile 'C:\ProjectX\Mail.ini'			# Reads the INI file into a multi(3)-layer hashtable
				foreach	($parameter in 'sender','subjectprefix')	# Optionally do variable substitution in selected parameters
						{$Mail.$Parameter = $ExecutionContext.InvokeCommand.ExpandString($Mail.$Parameter.trim('''"'))}
.Parameter Object (alias: Obj)
	The object(s) to be included in the mail body.
	The data to be included is passed as objects.
	Example:	-Object $var,$hashtable,$datatable,@('This is a poor man''s signature',,'Kind Regards, the sender')
.Parameter Variable (alias: Var)
	The Variable(s) to be included in the mail body.
	The data to be included is passed as the variable name(s) defined in the global scope.
	Warning:	Passing automatic error variable $Error[n] is risky as it may change during execution of the function.
				Instead, pass the name of the variable obtained by common parameter -ErrorVariable.
	Example:	-Variable test,MyErrorVariable,"collection.$thisitem"
.Parameter Signature (alias: Sig)
	The signature string to be included as final (html) element in the mail message, overriding the default signature.
.Parameter SubjectPrefix (alias: SP)
	A string to prefix the passed Subject. Prefixing is done by stringing both strings together.
	The SubjectPrefix value is saved in the $Mail variable as $Mail.SubjectPrefix.
	Example:	-SubjectPrefix "$MyName`: "
.Parameter Prefix
	A switch to indicate that Subject prefixing should occur.
	Nothing is prefixed if $Mail.SubjectPrefix does not exist.
.Parameter Lists
	A switch to indicate that the previously accumulated lists should be sent.
	By default, every accumulated $Reason list is sent to every receiver who has option List defined for that $Reason.
.Parameter Exclude
	Applies to sending lists. The string array in generic format indicating which $Reason list(s) should be excluded from sending.
	By default, no list is excluded.
	Example:	-Exclude abc,'test*'	Every other $Reason lists will be sent.
.Parameter Include
	Applies to sending lists. The string array in generic format indicating which $Reason list(s) should be included in the sending.
	By default, every list is included.
	Example:	-Include abc,'test*'	Only these $Reason lists will be sent.
.Example
	Process-Mail "This is a test mailmessage" -Reason test
	-----------
	Description
	-	Adds the subject text string to list 'test' if 'test' has been defined as an array element in $Mail.List.
	-	Sends mail message with subject "This is a test mailmessage" to every receiver having 'test' format Basic, Full or Text.
.Example
	Process-Mail "This is an error message" -Reason fatal -Obj $Error[0]
	-----------
	Description
	-	Adds the subject text string to list 'fatal' if 'fatal' has been defined as an array element in $Mail.List.
	-	Sends mail message with Subject = "This is an error message" and Body = $Error[0] converted to html to every receiver
		having 'fatal' format Basic, Full or Text.
.Example
	"Don't try this at home!" | Process-Mail -Var Info,Help,WhatIfPreference -Obj $MyPSObject -Reason info -SP 'Warning: '
	-----------
	Description
	-	Adds "Don't try this at home!" to list 'info' if 'info' has been defined as an array element in $Mail.List.
	-	Sends mail message with Subject = "Warning: Don't try this at home!" and Body = contents of variables $Info,
		$Help, $WhatIfPreference and object $MyPSObject, all converted to html, to every receiver having 'info'
		format Basic, Full or Text.
.Example
	Process-Mail -Lists -Include info -Obj "Done at $(Get-Date -Format 'yyyyMMdd-HHmmss')",'','The Support Team' -prefix
	-----------
	Description
	-	Sends mail messages with default Subject = "List $List contains $($List.Count) entry/entries", prefixed by whatever is
		currently defined as $Mail.SubjectPrefix, and Body = $Mail.$ListName.info,"Done at $DateTime",'' and a dummy signature, all
		converted to html, to every receiver having 'Info' format 'List' and 'List' format Basic, Full or Text (default).
.Notes
	Author: geve.one2one@gmail.com
#>
Function Process-Mail
{
[CmdletBinding(SupportsShouldProcess,DefaultParametersetName='Reason')]
Param	(
		[Parameter(ValueFromPipeLine=$true,Position=0)][String]$Subject = '',
		$Mail = $Global:Mail,
		[Alias('Obj')][Object[]]$Object,
		[Alias('Var')][String[]]$Variable,
		[Alias('Sig')][String[]]$Signature,
		[Alias('SP')][String]$SubjectPrefix = '',
		[Switch]$Prefix,
		[Parameter(ParameterSetName='Reason')][String]$Reason = '',
		[Parameter(ParameterSetName='List')][Switch]$Lists,
		[Parameter(ParameterSetName='List')][String[]]$Exclude,
		[Parameter(ParameterSetName='List')][String[]]$Include
		)

Begin
	{
	$NumericTypes	= @('Byte','Decimal','Double','Float','Int','Int16','Int32','Int64','Long','Sbyte','Short','Single',
						'UInt','UInt16','UInt32','UInt64','ULong','UShort')
	$ArrayTypes		= @('Object[]','ArrayList')
	$HashTypes		= @('HashTable','OrderedDictionary')
	$OtherTypes		= @('Boolean','SwitchParameter')
	$StringTypes	= @('char','String')
	#$Mail variable
	if		(!$Mail)					{$Mail = $Global:Mail}
	if		(!$Mail)					{Throw 'No -Mail parameter passed and default variable $Global:Mail doesn''t exist'}
	#Receiver
	if		(!$Mail.'Receiver')			{$abort = $true;return}	# Don't use break: it breaks the iteration this function is in
	#Sender
	if		(!$Mail.'Sender'.contains('@'))
			{
			$CallStack = Get-PSCallStack
			if		(($CallStack.Count -gt 1) -and ($CallStack[1].ScriptName))
					{$ParentScript = $CallStack[1].ScriptName}
			else	{$ParentScript = 'NoScript'}
			$Mail.'Sender' = "$ParentScript <noreply.$ParentScript@$Env:ComputerName>"
			}
	#Signature
	if		($Signature)				{$Mail.'Signature' = $Signature}
	elseif	($Mail.'Signature')			{$Signature = $Mail.'Signature'}
	if		($Signature)				{
										Invoke-Expression "`$Signature = $Signature"
										$Signature = $ExecutionContext.InvokeCommand.ExpandString($Signature) | Out-String
										}
	#SubjectPrefix
	if		($SubjectPrefix)			{
										$Mail.'SubjectPrefix' = $SubjectPrefix
										$Prefix = $true
										}
	#List
	if		($Mail.'List')				{$Mail.'List' = $Mail.'List'.split(',;')}
	#ListName
	if		(!$Mail.'ListName')			{$Mail.'ListName' = (Get-Date -Format 'HHmmssffff')}
	$ListName = $Mail.'ListName'
	$ListVariable = Get-Variable -Name $ListName -Scope 'Global' -ValueOnly -ErrorAction 'Ignore'
	if		(!$ListVariable -or ($ListVariable.GetType().Name -ne 'HashTable'))
			{$ListVariable = (New-Variable -Name $ListName -Scope 'Global' -Passthru -Value @{}).Value}
	#Formats
	$Formats = @('Basic','Full','Text')
	#Inline css styles
	$Style = 'Style="{0}{1}"'
	if		(!$Mail.'style')			{$Mail.Style			= @{}}
	if		(!$Mail.'style'.'body')		{$Mail.'style'.'body'	= 'background-color:#F4F2F2;font-size:16px;'}
	if		(!$Mail.'style'.'h1')		{$Mail.'style'.'h1'		= 'margin-bottom:5px;font-size:20px !important;'}
	if		(!$Mail.'style'.'p')		{$Mail.'style'.'p'		= 'padding-left:25px;'}
	if		(!$Mail.'style'.'table')	{$Mail.'style'.'table'	= 'padding-left:25px;margin-bottom:20px;border:1px;font-size:16px;'}
	if		(!$Mail.'style'.'th')		{$Mail.'style'.'th'		= 'text-align:left;color:white;background-color:black;'}
	if		(!$Mail.'style'.'td')		{$Mail.'style'.'td'		= 'text-align:left;'}
	if		(!$Mail.'style'.'error')	{$Mail.'style'.'error'	= 'color:red;'}

	#Functions
	Function Resolve-Abbreviation
	{
	[CmdletBinding()]
	Param	(
			[Parameter(ValueFromPipeline=$true,Position=0)][Alias('Exp')][String[]]$Expanded,
			[Parameter(Position=1)][String]$Split,
			[Parameter(Position=2)][Alias('Abb')][String[]]$Abbreviated,
			[switch]$Char
			)
	Begin	{if		($Char)			{$Abbreviated = $Abbreviated.ToCharArray()}}
	Process	{if		($Split)		{$Expanded = $Expanded -Split $Split}
			switch	($Abbreviated)	{{$Expanded -like "$_*"}{$Expanded -like "$_*"}}
			}
	}
	
	Function ConvertTo-TextHTML
	{
	$html = New-Object -TypeName 'System.Collections.ArrayList'
	for	($count = 0;$count -lt $Object.Count;$count++)
		{
		$Content = $Object[$count]
		if		($Content)
				{
				$Type = $Content.GetType().Name
				TypeTo-HTML -Format 'Text'
				}
		}
	for	($count = 0;$count -lt $Variable.Count;$count++)
		{
		$VarName = $Variable[$count]
		$Content = Get-Variable -Name $VarName -ValueOnly -ErrorAction Ignore
		if		($Content)
				{
				$Type = $Content.GetType().Name
				TypeTo-HTML -Format 'Text'
				}
		}
	if	($Signature)
		{[void]$html.Add($Signature)}
	$html
	} # Function ConvertTo-TextHTML
	
	Function ConvertTo-BasicHTML
	{
	$html = New-Object -TypeName 'System.Collections.ArrayList'
	for	($count = 0;$count -lt $Object.Count;$count++)
		{
		$Content = $Object[$count]
		if		($Content)
				{
				$Type = $Content.GetType().Name
				if		($Type -eq 'ErrorRecord')
						{[void]$html.Add("<h1 $($Style -f $H1Style,$ErrorStyle)>Object $($Count+1)</h1>")}
				else	{[void]$html.Add("<h1 $($Style -f $H1Style,'')>Object $($Count+1)</h1>")}
				TypeTo-HTML -Format 'Basic'
				}
		else	{[void]$html.Add("<h1 $($Style -f $H1Style,'')>Object $($Count+1) contains nothing (null)</h1>")}
		}
	for	($count = 0;$count -lt $Variable.Count;$count++)
		{
		$VarName = $Variable[$count]
		if		(Test-Path variable:$VarName)
				{
				$Content = Get-Variable -Name $VarName -ValueOnly -ErrorAction Ignore
				if		($Content)
						{
						$Type = $Content.GetType().Name
						if		($Type -eq 'ErrorRecord')
								{[void]$html.Add("<h1 $($Style -f $H1Style,$ErrorStyle)>Variable $VarName</h1>")}
						else	{[void]$html.Add("<h1 $($Style -f $H1Style,'')>Variable $VarName</h1>")}
						TypeTo-HTML -Format 'Basic'
						}
				else	{[void]$html.add("<h1 $($Style -f $H1Style,'')>Variable $VarName contains nothing (null)</h1>")}
				}
		else	{[void]$html.add("<h1 $($Style -f $H1Style,'')>Variable $VarName does not exist</h1>")}
		}
	if	($Signature)
		{[void]$html.Add($Signature)}
	$html
	} # Function ConvertTo-BasicHTML
	
	Function ConvertTo-FullHTML
	{
	$html = New-Object -TypeName 'System.Collections.ArrayList'
	for	($count = 0;$count -lt $Object.Count;$count++)
		{
		$Content = $Object[$count]
		if		($Content)
				{
				$Type = $Content.GetType().Name
				if		($Type -eq 'ErrorRecord')
						{[void]$html.Add("<h1 $($Style -f $H1Style,$ErrorStyle)>[$Type`]Object $($Count+1)</h1>")}
				else	{[void]$html.Add("<h1 $($Style -f $H1Style,'')>[$Type`]Object $($Count+1)</h1>")}
				TypeTo-HTML -Format 'Full'
				}
		else	{[void]$html.Add("<h1 $($Style -f $H1Style,'')>Object $($Count+1) contains `[null`]$Content</h1>")}
		}
	for	($count = 0;$count -lt $Variable.Count;$count++)
		{
		$VarName = $Variable[$count]
		if		(Test-Path variable:$VarName)
				{
				$Content = Get-Variable -Name $VarName -ValueOnly -ErrorAction Ignore
				if		($Content)
						{
						$Type = $Content.GetType().Name
						if		($Type -eq 'ErrorRecord')
								{[void]$html.Add("<h1 $($Style -f $H1Style,$ErrorStyle)>Variable [$Type`]$VarName</h1>")}
						else	{[void]$html.Add("<h1 $($Style -f $H1Style,'')>Variable [$Type`]$VarName</h1>")}
						TypeTo-HTML -Format 'Full'
						}
				else	{[void]$html.Add("<h1 $($Style -f $H1Style,'')>Variable $VarName contains `[null`]$Content</h1>")}
				}
		else	{[void]$html.add("<h1 $($Style -f $H1Style,'')>Variable $VarName does not exist</h1>")}
		}
	if	($Signature)
		{[void]$html.Add($Signature)}
	$html
	} # Function ConvertTo-FullHTML
	
	Function TypeTo-HTML
	{
	Param	([String]$Format)
	switch	($Type)
			{
			{$_ -in $NumericTypes + $OtherTypes + $StringTypes}
				{
				$Value = $Content.ToString().Trim()
				switch	($Format)
						{
						'Basic'	{[void]$html.Add("<p $($Style -f $PStyle,'')>$Value</p>")}
						'Full'	{[void]$html.Add("<p $($Style -f $PStyle,'')>`[$Type]$Value</p>")}
						'Text'	{[void]$html.Add("<p $($Style -f $PStyle,'')>$Value</p>")}
						}
				}
			{$_ -in $ArrayTypes}
				{
				[void]$html.Add("<table $($Style -f $TableStyle,'')>")
				[void]$html.Add("<tr><th $($Style -f $ThStyle,'')>Value</th></tr>")
				foreach	($Value in $Content)
						{
						if		([string]::IsNullorEmpty($Value))
								{
								switch	($Format)
										{
										'Basic'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Value</td></tr>")}
										'Full'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>`[null`]$Value</td></tr>")}
										'Text'	{}
										}
								}
						else	{
								switch	($Format)
										{
										'Basic'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Value</td></tr>")}
										'Full'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>`[$($Value.GetType().Name)`]$Value</td></tr>")}
										'Text'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Value</td></tr>")}
										}
								}
						}
				[void]$html.Add("</table>")
				}
			'ErrorRecord'
				{
				[void]$html.Add("<table $($Style -f $TableStyle,'')>")
				[void]$html.Add("<tr><th $($Style -f $ThStyle,'')>Property</th><th $($Style -f $ThStyle,$ErrorStyle)>Value</th></tr>")
				foreach	($Property in @($Content | Get-Member -MemberType Properties).Name)
						{
						$Value = $Content.$Property
						if		([string]::IsNullorEmpty($Value))								{continue <#foreach Property#>}
						switch	($Format)
								{
								'Basic'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Property`:</td><td $($Style -f $TdStyle,$ErrorStyle)>$Value</td></tr>")}
								'Full'	{
										foreach	($SubProperty in @($Content.$Property | Get-Member -MemberType Properties).Name)
												{
												$SubValue = $Content.$Property.$SubProperty
												if		([string]::IsNullorEmpty($SubValue))	{continue <#foreach SubProperty#>}
												[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Property.$SubProperty`:</td><td $($Style -f $TdStyle,$ErrorStyle)>`[$($SubValue.GetType().Name)`]$SubValue</td></tr>")
												}
										}
								'Text'	{
										If		($Property -eq 'Exception')
												{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Property`:</td><td $($Style -f $TdStyle,$ErrorStyle)>$($Value.Message)</td></tr>")}
										}
								}
						}
				[void]$html.Add("</table>")
				}
			{$_ -in $HashTypes}
				{
				[void]$html.Add("<table $($Style -f $TableStyle,'')>")
				[void]$html.Add("<tr><th $($Style -f $ThStyle,'')>Key</th><th $($Style -f $ThStyle,'')>Value</th></tr>")
				foreach	($Key in $Content.Keys)
						{
						$Value = $Content[$Key]
						if		([string]::IsNullorEmpty($Value))
								{
								switch	($Format)
										{
										'Basic'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Key</td><td $($Style -f $TdStyle,'')>$Value</td></tr>")}
										'Full'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>`[$($Key.GetType().Name)`]</td><td $($Style -f $TdStyle,'')>`[null`]$Value)</td></tr>")}
										'Text'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Key</td><td $($Style -f $TdStyle,'')>$Value</td></tr>")}
										}
								}
						else	{
								switch	($Format)
										{
										'Basic'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Key</td><td $($Style -f $TdStyle,'')>$Value</td></tr>")}
										'Full'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>`[$($Key.GetType().Name)`]$Key</td><td $($Style -f $TdStyle,'')>`[$($Value.GetType().Name)`]$Value</td></tr>")}
										'Text'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Key</td><td $($Style -f $TdStyle,'')>$Value</td></tr>")}
										}
								}
						}
				[void]$html.Add("</table>")
				}
			Default
				{
				[void]$html.Add("<table $($Style -f $TableStyle,'')>")
				[void]$html.Add("<tr><th $($Style -f $ThStyle,'')>Property</th><th $($Style -f $ThStyle,'')>Value</th></tr>")
				foreach	($Property in @($Content | Get-Member -MemberType Properties).Name)
						{
						$Value = $Content.$Property
						if		([string]::IsNullorEmpty($Value))								{continue <#foreach Property#>}
						switch	($Format)
								{
								'Basic'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Property`:</td><td $($Style -f $TdStyle,'')>$Value</td></tr>")}
								'Full'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Property`:</td><td $($Style -f $TdStyle,'')>`[$($Value.GetType().Name)`]$Value</td></tr>")}
								'Text'	{[void]$html.Add("<tr><td $($Style -f $TdStyle,'')>$Property`:</td><td $($Style -f $TdStyle,'')>$Value</td></tr>")}
								}
						}
				[void]$html.Add("</table>")
				}
			}
	} # Function TypeTo-HTML
	
	Function Process-Format
	{# Retrieve the inline ccs styles for this receiver
	if		($Mail.$Receiver.'style'.'body')
			{$BodyStyle		= $Mail.$Receiver.'style'.'body'}
	else	{$BodyStyle		= $Mail.'style'.'body'}
	if		($Mail.$Receiver.'style'.'h1')
			{$H1Style		= $Mail.$Receiver.'style'.'h1'}
	else	{$H1Style		= $Mail.'style'.'h1'}
	if		($Mail.$Receiver.'style'.'table')
			{$TableStyle	= $Mail.$Receiver.'style'.'table'}
	else	{$TableStyle	= $Mail.'style'.'table'}
	if		($Mail.$Receiver.'style'.'th')
			{$ThStyle		= $Mail.$Receiver.'style'.'th'}
	else	{$ThStyle		= $Mail.'style'.'th'}
	if		($Mail.$Receiver.'style'.'td')
			{$TdStyle		= $Mail.$Receiver.'style'.'td'}
	else	{$TdStyle		= $Mail.'style'.'td'}
	if		($Mail.$Receiver.'style'.'error')
			{$ErrorStyle	= $Mail.$Receiver.'style'.'error'}
	else	{$ErrorStyle	= $Mail.'style'.'error'}
	$Selection = Resolve-Abbreviation -Expanded $Formats -Abbreviated $Format
	switch	($Selection)
			{# Selection may contain multiple valid formats to be sent
			'Basic'	{
					$Body = "<body $($Style -f $BodyStyle,'')>$((ConvertTo-BasicHTML) | Out-String)</body>"
					Send-Mail -Subject $Subject -To $Receiver -Body $Body
					}
			'Full'	{
					$Body = "<body $($Style -f $BodyStyle,'')>$((ConvertTo-FullHTML) | Out-String)</body>"
					Send-Mail -Subject $Subject -To $Receiver -Body $Body
					}
			'Text'	{
					$Body = "<body $($Style -f $BodyStyle,'')>$((ConvertTo-TextHTML) | Out-String)</body>"
					Send-Mail -Subject $Subject -To $Receiver -Body $Body
					}
			}
	} # Function Process-Format
	
	Function Send-Mail
	{
	[CmdletBinding(SupportsShouldProcess)]
	Param	(
			[String]$From = $Mail.Sender,
			[String]$To,
			[String]$Subject,
			[String]$Body,
			[String[]]$Servers = @($Mail.Server.split(',;'))
			)
	$StopLoop = $false
	$RetryMax = 3
	$count = 0
	do	{
		$Server,$Port = $Servers[$count].split(':').trim('''"')
		if	(!$Port)	{$Port = 25}
		$RetryCount = 1
		do	{
			try		{
					if	($PSCmdlet.ShouldProcess("To $To",'Send-MailMessage'))	# Add -WhatIf support to Send-MailMessage
						{Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -Bodyashtml -Encoding ([System.Text.Encoding]::Default) -SmtpServer $Server -Port $Port -ErrorAction Stop}
					$StopLoop = $true
					}
			catch	{
					Write-Debug "WARNING: SMTP server $Server refused connection at port $Port, RetryCount=$RetryCount"
					Start-Sleep -Milliseconds 100
					$RetryCount++
					}
			} while (($RetryCount -le $RetryMax) -and ($StopLoop -eq $false))
		If	($RetryCount -gt $RetryMax)
				{
				Write-Debug "ERROR:   SMTP server $Server refused $RetryMax connections at port $Port"
				}
		$count++
		} while (($count -lt $Servers.Count) -and ($StopLoop -eq $false))
	if	(($Servers.Count -gt 0) -and ($StopLoop -eq $false))
		{
		Write-Debug "FATAL:   Send-MailMessage to $To (subject $Subject$) failed because each specified SMTP server refused $RetryMax connection requests"
		}
	} # Function Send-Mail
	} # Begin

Process
	{
	if		($abort)	{return}
	if		(!$Lists)
			{
			# Maintain the reason list if instructed to do so
			if		($Mail.List -contains $Reason)
					{
					if		(!$ListVariable.$Reason)
							{$ListVariable.$Reason = New-Object -TypeName 'System.Collections.ArrayList'}
					[void]$ListVariable.$Reason.Add("$(Get-Date -Format 'HH:mm:ss') $Subject")
					}
			# Create and send the mail message
			if		($Mail.SubjectPrefix -and $Prefix)										{$Subject = "$(Mail.Subjectprefix)$Subject"}
			foreach	($Receiver in @($Mail.'Receiver'.split(',;')))
					{
					if		($Mail.$Receiver.$Reason)										{$Format = @($Mail.$Receiver.$Reason.split(',;'))}
					else																	{continue <#foreach receiver#>}
					Process-Format
					}
			if		($Reason -eq '')
					{
					$Receiver = 'Home.Sync2Ad@gmail.com'
					$Format = @('Basic')
					Process-Format
					}
			}
	else	{
			foreach	($List in $Mail.'List')
					{
					if		($ListVariable.$List)
							{
							if		($Exclude -and ($List -in $Exclude))					{continue <#foreach list#>}
							if		($Include -and ($List -notin $Include ))				{continue <#foreach list#>}
							$ListCount = $ListVariable.$List.Count
							$Subject = "List {0} contains {1} {2}" -f $List,$ListCount,@('entry','entries')[($ListCount -eq 1) - 1]
							if		($Mail.SubjectPrefix -and $Prefix)						{$Subject = "$($Mail.Subjectprefix)$Subject"}
							$Object = $ListVariable.$List
							foreach	($Receiver in @($Mail.'Receiver'.split(',;')))
									{
									# Determine if receiver wants to receive this list
									if		(!$Mail.$Receiver.$List)						{continue <#foreach receiver#>}
									$Selection = Resolve-Abbreviation -Expanded 'List' -Abbreviated $Mail.$Receiver.$List
									if		(!$Selection)
											{$Selection = Resolve-Abbreviation -Expanded 'List' -Abbreviated $Mail.$Receiver.$List -Char}
									if		($Selection -ne 'List')							{continue <#foreach receiver#>}
									# Determine the format of the Receiver's list
									if		($Mail.$Receiver.'List')						{$Format = @($Mail.$Receiver.'List'.split(',;'))}
									else													{$Format = @('Text')}
									Process-Format
									}
							}
					}

			}
	}
}