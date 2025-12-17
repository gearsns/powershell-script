#### 設定
$site = "サイト名"
$tenantname = "テナント名"
$clientid = "クライアントID"
$tenant = "$tenantname.onmicrosoft.com"
$thumbprint = "サムプリント"
$fromListName = "取得もとの一覧の名前"
# 適宜修正
$siteUrl = "https://$tenantname.sharepoint.com/sites/$site"
####

$folderName = [regex]::Replace($fromListName, "[\[\(:].*$", "")
# 接続
Connect-PnPOnline -Url $siteUrl -clientid $clientid -Tenant $tenant -Thumbprint $thumbprint
# データ取得
$PageItems = Get-Content -Path "$folderName/$($folderName)_コメントあり.json" | ConvertFrom-Json

# 先リストの情報を保存
$listFields = @()
$fields = Get-PnPField -List "SitePages"
foreach ($field in $fields) {  
	$listFields += New-Object PSObject -Property ([ordered]@{
							"Title"            = $field.Title
							"Type"             = $field.TypeAsString
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
$listFields | ConvertTo-Json -Depth 10 | Out-File -FilePath "$folderName/ToFields.json"
#exit
foreach($item in $PageItems){
	$pageName = "$($fromListName)_$($item.ID)"
	$Modified = $item.Modified
	if ($item.Comments){
		foreach($comment in $($item.comments | Sort-Object -Property ListItemId)){
			if ($Modified -lt $comment.CreatedDate){
				$Modified = $comment.CreatedDate
			}
		}
	}
	Write-Host $pageName
	# ページ追加
	try {
		$page = Get-PnPPage -Identity $pageName
		$pageItem = Get-PnPListItem -List "SitePages" -Id  $page.pageId
		if ($Modified -gt $pageItem.FieldValues.Modified){
			Write-Host "ページ更新：$($Modified):$($pageItem.Modified):$pageName ($($item.Title))"
		} else {
			continue
		}
		Remove-PnPPage -Identity $pageName
	} catch {
		#Write-Host $_
		Write-Host "ページ追加：$pageName ($($item.Title))"
	}
	$page = Add-PnPPage -Name $pageName -Layout Article -PromoteAs NewsArticle
	# ページのタイトル設定
	Write-Host "  タイトル設定"
	$p = Set-PnPPage -Identity $pageName -Title $item.Title
	# 本文設定
	Write-Host "  本文設定"
	# HTMLをテキストに変換(デザインもそのまま使用するのであれば、$item.new_contextをそのまま Add-PnPPageTextPart に)
	$HtmlParser = [AngleSharp.Html.Parser.HtmlParser]::new()
	$HtmlDocument = $HtmlParser.ParseDocument($item.new_context)

	$walker = $HtmlDocument.CreateTreeWalker($HtmlDocument.Body)
	$text = ""
	while($walker.ToNext()){
		$current = $walker.Current
		if ($current.NodeType -eq [AngleSharp.Dom.NodeType]::Text) {
			$text += $current.TextContent
		}
		# 要素ノード (<div>や<p>など) の場合
		elseif ($current.LocalName -eq "p" -or $current.LocalName -eq "div" -or $current.LocalName -eq "br") {
			$text += "`n"
		}
	}
	$t = "<p>$([regex]::Replace($text, "`n", "</p>`n<p>"))</p>"
	#  添付ファイル(=イメージファイル)
	if (Test-Path "$folderName/att/$($item.ID)"){
		$attFiles = Get-ChildItem "$folderName/att/$($item.ID)" -File
		foreach($file in $attFiles){
			if ($file.Name -match "\.(jpeg|jpg|png)"){
				Add-PnPFile -Path $file.FullName -Folder "SiteAssets/SitePages/$pageName"
				$imageUrl = "/sites/$site/SiteAssets/SitePages/$pageName/$($file.Name)"
				$t += "<hr><img src='$imageUrl' style='max-width:50%'>"
			}
		}
	}
	$p = Add-PnPPageTextPart -Page $pageName -Text $t
	# コメント
	if ($item.Comments){
		$order = 2
		Write-Host "  コメント設定"
		$comments = "<hr><h4>コメント</h4><ul>"
		foreach($comment in $($item.comments | Sort-Object -Property ListItemId)){
			$d = (Get-Date $comment.CreatedDate.ToString("yyyy-MM-ddTHH:mmZ")).ToString("yyyy-MM-dd HH:mm")
			$comments += "<li><dl><dt>$d : $($comment.UserName)($($comment.UserMail))<dd>"
			$comments += [regex]::Replace($comment.CommentText, "\n", "<br>")
			$comments += "</li>"
		}
		$comments += "</ul>"
		$SectionNumber = 1  # ページに存在するセクション番号
		$ColumnNumber = 1   # セクション内の列番号
		$p = Add-PnPPageTextPart -Page $pageName -Section $SectionNumber -Column $ColumnNumber -Text $comments -Order $order
	}
	# 公開
	Write-Host "  公開"
	$p = Set-PnPPage -Identity $pageName -ShowPublishDate $true -Publish
	# 所有者などの設定を変更
	Write-Host "  所有者などの設定を変更"
	$p = Set-PnPListItem -List "SitePages" -Identity $page.pageId `
		-Values @{
			'Author' = $item.Author.Email
			'Editor' = $item.Author.Email
			'Created' = $item.Created
			'Modified' = $item.Modified
			'FirstPublishedDate' = $item.PublishingStartDate
		}
	Write-Host "  反映"
	$p = Set-PnPPage -Identity $pageName -ShowPublishDate $true -Publish
	$p = Set-PnPListItem -List "SitePages" -Identity $page.pageId `
		-Values @{
			'Author' = $item.Author.Email
			'Editor' = $item.Author.Email
			'Created' = $item.Created
			'Modified' = $item.Modified
			'FirstPublishedDate' = $item.PublishingStartDate
		}
}

DisConnect-PnPOnline
