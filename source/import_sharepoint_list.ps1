#### 設定
$site = "サイト名"
$tenantname = "テナント名"
$clientid = "クライアントID"
$tenant = "$tenantname.onmicrosoft.com"
$thumbprint = "サムプリント"
$listName = "一覧の名前"
$listID = "リストのID" # 画像とかの保存場所
$fromListName = "取得もとの一覧の名前"
$targetDate = "2025-12-01" # 対象日付以降のデータを取り込む
$defUserAddress = "xxxx@xxxx" # IDが削除されている場合の代替ユーザ
# 適宜修正
$siteUrl = "https://$tenantname.sharepoint.com/sites/$site"
$newurl = "https://$tenantname.sharepoint.com/sites/$site/SiteAssets/Lists/$listID/"
$base_newurl = "https://$tenantname.sharepoint.com/sites/xxxxx_new_site_xxxxx/"
####

$folderName = [regex]::Replace($fromListName, "[\[\(:].*$", "")

$PageItems = Get-Content -Path "$folderName/$($folderName)_コメントあり.json" | ConvertFrom-Json
Connect-PnPOnline -Url $siteUrl -clientid $clientid -Tenant $tenant -Thumbprint $thumbprint

# データを更新する際には、作成日で特定する
# ※途中で件名が変わる場合もあるため
# 厳密にするのであればIDを保持？
$title_list = @{}
$TargetPageItems = Get-PnPListItem -List $listName -Fields Created,ID,Category,Title, PublishingStartDate,PublishingEndDate,ExpiredDate, Body, OriginatorId, Author, Attachments,LikedBy,Editor -PageSize 1000
$TargetPageItems = $TargetPageItems.FieldValues | ConvertTo-Json -Depth 10 | ConvertFrom-Json
foreach($item in $TargetPageItems){
	$orgItem = $title_list[$item.Created]
	if ($orgItem){
		Write-Host "Dup:$($item.Created):$($item.Title) == $($orgItem.Created):$($orgItem.Title)"
	}
	$title_list[$item.Created] = $item
}

$targetItemDate = Get-Date $targetDate

$global:email_list = @{}

# メールアドレスの取得
function Get-UserEmailCached {
[CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $user
    )
	$email = $user.Email
	if (-not $email){
		if ($global:email_list.ContainsKey($user.LookupId)){
			return $global:email_list[$user.LookupId]
		}
		Write-Host "$($user.LookupId):$($user.LookupValue)"
		try {
			$users = Get-PnPUser -Identity $user.LookupId
			if ($users.Count -gt 0) {
				$global:email_list[$users[0].Email] = $users[0].Email
				$global:email_list[$user.LookupId] = $users[0].Email
				return $email
			} else {
				$global:email_list[$user.LookupId] = $defUserAddress
				return $defUserAddress
			}
		} catch {
			$global:email_list[$user.LookupId] = $defUserAddress
			return $defUserAddress
		}
	}
	if ($global:email_list.ContainsKey($email)){
		return $global:email_list[$email]
	}
	try {
		$user = Get-PnPUser -Identity $user.Email
		if ($user.Count -gt 0) {
			$global:email_list[$email] = $email
			$global:email_list[$user.LookupId] = $email
			return $email
		} else {
			$global:email_list[$email] = $defUserAddress
			return $defUserAddress
		}
	} catch {
		$global:email_list[$email] = $defUserAddress
		return $defUserAddress
	}
}
foreach($item in $PageItems){
	$author = Get-UserEmailCached -user $item.Author
	$editor = Get-UserEmailCached -user $item.Editor
	# リンクを付け替え
	$textContent = [regex]::Replace($item.new_context, $base_newurl, "$newurl")
	# 追加
	$itemValues = @{
		"Title" = $item.Title;
		"Category" = $item.Category;
		"PublishingStartDate" = $item.PublishingStartDate;
		"PublishingEndDate" = $item.PublishingEndDate;
		"ExpiredDate" = $item.ExpiredDate;
		"Body" = $textContent;
		"Modified" = $item.Modified;
		"Created" = $item.Created;
		"Author" = $author;
		"Editor" = $editor;
	}
	# ここで必要なら項目をカスタマイズ

	#
	$orgItem = $title_list[$item.Created]
	if ($orgItem){
		if ($item.Modified -ne $orgItem.Modified){
			Write-Host "Modified:$($item.Modified):$($orgItem.Modified):$($orgItem.Title)"
			Set-PnPListItem -List $listName -Identity $orgItem.Id -Values $itemValues
			$newItem = $orgItem
		} elseif ($item.Title -ne $orgItem.Title){
			Write-Host "Modified:$($item.Title):$($orgItem.Title):$($orgItem.Title)"
			continue
		} else {
			continue
		}
	} elseif ($item.Created -gt $targetItemDate){
		Write-Host "Add:$($item.Created):$($item.Title)"
		$newItem = Add-PnPListItem -List $listName -Values $itemValues
	} else {
		continue
	}
	$pageName = "$($fromListName)_$($item.ID)"
	Write-Host "リスト追加：$pageName ($($item.Title))"
	Write-Host "  IDチェック"

	## 画像
	if (Test-Path "$folderName/image/$($item.ID)"){
		Write-Host "  画像"
		$imageFiles = Get-ChildItem "$folderName/image/$($item.ID)" -File
		foreach($file in $imageFiles){
			Add-PnPFile -Path $file.FullName -Folder "SiteAssets/Lists/$listID/$($item.ID)"
		}
	}
	# 添付ファイル
	if (Test-Path "$folderName/att/$($item.ID)"){
		Write-Host "  添付ファイル"
		$attFiles = Get-ChildItem "$folderName/att/$($item.ID)" -File
		foreach($file in $attFiles){
			Add-PnPListItemAttachment -List $listName -Identity $newItem.ID -Path $file.FullName
		}
	}
	Write-Host "  更新日付再設定"
	Set-PnPListItem -List $listName -Id $newItem.ID -Values @{
		"Modified"=$item.Modified;
		"Editor"=$editor;
	}
}

DisConnect-PnPOnline
