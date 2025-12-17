
#### 設定
$site = "サイト名"
$tenantname = "テナント名"
$clientid = "クライアントID"
$thumbprint = "サムプリント"
$listName = "一覧の名前"
# 適宜修正
$tenant = "$tenantname.onmicrosoft.com"
$base_oldurl = "https://$tenantname.sharepoint.com/sites/$site/RscLib/" 
$base_newurl = "https://$tenantname.sharepoint.com/sites/xxxxx_new_site_xxxxx/"
$siteUrl = "https://$tenantname.sharepoint.com/sites/$site"
$skipComment = $listName -ne "コメントのダウンロード？"
####

$folderName = [regex]::Replace($listName, "[\[\(:].*$", "")
New-Item  "$folderName" -ItemType Directory -ErrorAction SilentlyContinue

$savedPageItems = @()
if (Test-Path "$folderName/$($folderName)_コメントあり.json"){
	$savedPageItems = Get-Content -Path "$folderName/$($folderName)_コメントあり.json" | ConvertFrom-Json
}

# 接続
Connect-PnPOnline -Url $siteUrl -clientid $clientid -Tenant $tenant -Thumbprint $thumbprint
# 元リストの情報を保存
#  この情報に基づいて、必要なら一覧取得
#  Get-PnPListItemの-Fieldsパラメータを変更
$listFields = @()
$fields = Get-PnPField -List $listName
foreach ($field in $fields) {  
	$listFields += New-Object PSObject -Property ([ordered]@{
							"Title"            = $field.Title
							"Type"             = $field.TypeAsSstring
							"Internal Name"    = $field.InternalName
							"Static Name"      = $field.StaticName
							"Scope"            = $field.Scope
							"Type DisplayName" = $field.TypeDisplayName
							"Is read only?"    = $field.ReadOnlyField
							"Unique?"          = $field.EnforceUniqueValues
							"IsRequired"       = $field.Required
							"IsSortable"       = $field.Sortable
							"Schema XML"       = $field.SchemaXml
							"Description"      = $field.Description
							"Group Name"       = $field.Group
						})
}
$listFields | ConvertTo-Json -Depth 10 | Out-File -FilePath "$folderName/FromFields.json"
# 一覧取得
Write-Host "一覧取得"
$PageItems = Get-PnPListItem -List $listName -Fields ID,Category,Title, PublishingStartDate,PublishingEndDate,ExpiredDate, Body, OriginatorId, Author, Attachments,LikedBy,Editor -PageSize 1000
# 添付ファイルダウンロード
Write-Host "添付ファイルダウンロード"
foreach ($item in $PageItems.FieldValues) {
	$m = $savedPageItems | Where-Object { $_.ID -eq $item.ID }
	if ($m.Count -gt 0){
		if ($item.Attachments -eq $True) {
			$LastWriteTime = (Get-Item -Path "$folderName/att/$($item.ID)").LastWriteTime
			if ($item.Modified -gt $LastWriteTime){
				Write-Host "添付ファイルダウンロード?:$($item.ID):$($item.ExpiredDate):$($item.Title)"
				Rename-Item -Path "$folderName/att/$($item.ID)" -NewName "$($item.ID).old"
			} else {
				continue
			}
		} else {
			continue
		}
	}
	if ($item.Attachments -eq $True) {
		Write-Host "添付ファイルダウンロード:$($item.Title):$($item.ID)"
		New-Item  "$folderName/att/$($item.ID)" -ItemType Directory -ErrorAction SilentlyContinue
		$attachments = Get-PnPListItemAttachment -List $listName -Identity $item.ID -Path "$folderName/att/$($item.ID)"
		foreach ($att in $attachments) {
			Write-Host "  - $($att.FileName)"
		}
	}
}
# 一旦JSONに保存
$PageItems.FieldValues | ConvertTo-Json -Depth 10 | Out-File -FilePath "$folderName/$folderName.json"
$PageItems = Get-Content -Path "$folderName/$folderName.json" | ConvertFrom-Json
# 画像ファイルの保存
Write-Host "画像ファイルの保存"
foreach ($item in $PageItems) {
	$files = @()
	$baseurl = $siteUrl
	$oldurl  = $base_oldurl
	$newurl  = "$($base_newurl)/$($item.ID)/"
	$context = $item.Body
	$new_context = [regex]::Replace($context, "((?:href|src)\s*=\s*[""'])(.*?)([""'])", {
		param($m)
		try {
			$url = [System.Uri]::UnescapeDataString([System.Web.HttpUtility]::HtmlDecode($m.Groups[2].Value))
			$url = $url.Replace(":[A-Za-z]:/[A-Za-z]/", "")
			if ($url.StartsWith("//")) {
				$url = "https:$url"
			}
			elseif ($url.StartsWith("/")) {
				$url = "https://$tenantname.sharepoint.com$url"
			}
			elseif ($url.StartsWith(".")) {
				$url = "$baseurl$url"
			}
			elseif ($url.StartsWith("http")) {
				#Skip
			}
			else {
				$url = "$baseurl$url"
			}
			$uri = [System.Uri]$url
			$resolved = [System.Uri]::new($uri.AbsoluteUri, [System.UriKind]::Absolute)
			$url = [System.Uri]::UnescapeDataString($resolved.AbsoluteUri)
			if ($url.StartsWith($oldurl)){
				$global:files += $url
				$url = $url.Replace($oldurl, $newurl)
			}
			"$($m.Groups[1].Value)$url$($m.Groups[3].Value)"
		}
		catch {
			$m[0]
		}
	})
	if ($item.new_context){
		$item.new_context = $new_context
	} else {
		$item | Add-Member new_context $new_context
	}
	$m = $savedPageItems | Where-Object { $_.ID -eq $item.ID }
	if ($m.Count -gt 0){
		continue
	}
	if ($files.Count -gt 0){
		Write-Host "画像のダウンロード:$($item.ID):$($item.Title)"
		New-Item  "$folderName/image/$($item.ID)" -ItemType Directory -ErrorAction SilentlyContinue
		foreach($file in $files){
			Write-Host $file
			$filename = [regex]::Replace([regex]::Replace($file, ".*/", ""), "\?.*$", "")
			if (-not (Test-Path -Path "$folderName/image/$($item.ID)/$filename" -PathType Leaf)){
				Get-PnPFile -Url $file -AsFile -Path "$folderName/image/$($item.ID)" -Filename $filename
			}
		}
	}
}
# コメント取得
if ($skipComment){
	Write-Host "コメント取得はスキップ"
} else {
	Write-Host "コメント取得"
	foreach ($item in $PageItems) {
		Write-Host $item.ID
		$comments = Get-PnPListItemComment -List $listName -Identity $item.ID
		$allComments = @()
		if ($null -ne $comments) {
			foreach ($c in $comments) {
				$author = $c.Author
				Write-Host $c
				$allComments += [PSCustomObject]@{
					ListItemId = $c.Id
					CommentText = $c.Text
					UserMail = $author.Mail
					UserName = $author.Name
					CreatedDate = $c.CreatedDate
				}
			}
		}
		$allComments | ConvertTo-Json -Depth 5
		if ($item.Comments){
			$item.Comments = $allComments
		} else {
			$item | Add-Member Comments $allComments
		}
	}
}
$PageItems | Sort-Object -Property Modified | ConvertTo-Json -Depth 10 | Out-File -FilePath "$folderName/$($folderName)_コメントあり.json"
#
DisConnect-PnPOnline