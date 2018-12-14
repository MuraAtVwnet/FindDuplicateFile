<#
.SYNOPSIS
指定フォルダにある重複ファイルの検出、削除、移動をします。
比較は、ファイル名の「- コピー」「- Copy」「(数字)」を無視し、sha256 ハッシュと合わせて比較しますので、ファイル名が同じでも内容が異なる場合は別ファイル扱いにします。

重複したファイルは CSV 出力し、削除/移動した場合はリストにその操作も記録されます。
カレントディレクトリに実行ログも出力されます。
データ出力先は default カレントディレクトリですが、指定することも可能です(ファイル名指定はできません)

.DESCRIPTION
重複リストのみ出力(Remove/Moveオプションを指定していない時の動作)
    重複リスト出力だけをします

サブディレクトリも探査(-Recurse)
    サブディレクトリも探査します

削除(-Remove)
    重複したファイルを削除します
    重複リストには削除処理が記録されます

移動(-Move)
    重複したファイルを指定フォルダに移動します
    重複リストには移動処理と移動先ファイル名が記録されます
    移動先に同一ファイルがあった場合は「(数字)」をインクリメントします

ファイルパターン指定(-Pattern)
    特定拡張子のファイルだけ重複確認する場合は、パターン(*.jpg とか)を指定します
    複数指定する場合は、カンマ「,」で区切ってください

ショートカット作成(-CreateShortcut)
    ファイルを削除/移動する際にオリジナルファイルへのショートカットを残します

テスト(-WhatIf)
    実際の削除/移動はせず、動作確認だけします

.EXAMPLE
PS C:\Photo> .\FindDuplicateFile.ps1 -Recurse
カレントディレクトリ(C:\Photo)以下にある全ファイルの重複リストを出力します

.EXAMPLE
PS C:\Photo> .\FindDuplicateFile.ps1 -Recurse -Pattern *.jpg, *.png
カレントディレクトリ(C:\Photo)以下にある「*.jpg」と「*.png」の重複リストを出力します

動画や ISO などサイズの大きなファイルは、ハッシュ値の取得に時間がかかるので、特定種別のファイルだけ重複排除する場合はファイルパターンを指定するがお勧めです。

.EXAMPLE
PS C:\Photo> .\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複リストを出力します

.EXAMPLE
PS C:\Photo> .\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Remove
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを削除します

.EXAMPLE
PS C:\Photo> .\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Move C:\Photo\Backup
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを「C:\Photo\Backup」へ移動します

.EXAMPLE
PS C:\Photo> .\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Move C:\Photo\Backup -WhatIf
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを「C:\Photo\Backup」へ移動した場合の動作を確認します(実際の移動はしません)

.EXAMPLE
PS C:\Photo> .\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Move C:\Photo\Backup -CSVPath C:\Photo\CSV
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを「C:\Photo\Backup」へ移動します
重複リストを「C:\Photo\CSV」に出力します

.EXAMPLE
PS C:\Photo> .\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Move C:\Photo\Backup -CSVPath C:\Photo\CSV -LogPath C:\Photo\Log
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを「C:\Photo\Backup」へ移動します
重複リストは「C:\Photo\CSV」に、実行ログは「C:\Photo\Log」に出力します

.EXAMPLE
PS C:\Photo> .\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Move C:\Photo\Backup -CreateShortcut
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを「C:\Photo\Backup」へ移動し、代わりにショートカットを置きます
重複リストを「C:\Photo\CSV」に出力します

.PARAMETER Path
探索 Path
省略時はカレントディレクトリ以下を探索します
複数指定する場合はカンマで区切ります

.PARAMETER Pattern
対象ファイルパターン
省略時はすべてのファイルを対象にします
複数指定する場合はカンマで区切ります

.PARAMETER Recurse
サブディレクトリも探査します

.PARAMETER Remove
重複ファイルを削除します
Remove/Move が指定されていない場合は重複リストのみを出力します

.PARAMETER CreateShortcut
重複ファイルを削除/移動した時に、オリジナルファイルへのショートカットを残します

.PARAMETER CSVPath
CSV の出力先
省略時はカレントディレクトリに出力します

.PARAMETER CSVPath
CSV の出力先
省略時はカレントディレクトリに出力します

.PARAMETER LogPath
実行ログの出力先
省略時はカレントディレクトリに出力します

.PARAMETER AllList
重複確認する全ファイル リストを CSV 出力します

.PARAMETER WhatIf
実際の削除/移動はせず、動作確認だけします

.LINK
重複したファイル排除するスクリプト(PowerShell)
http://www.vwnet.jp/Windows/PowerShell/2018111601/FindDuplicateFile.htm
#>

###################################################
# 重複ファイル検出
###################################################
Param(
	[string[]]$Path,			# 探査する Path
	[switch]$Recurse,			# サブディレクトリも探査する
	[string[]]$Pattern,			# ファイルパターン
	[switch]$Remove,			# 削除実行
	[string]$Move,				# Move 先フォルダ
	[switch]$CreateShortcut,	# ショートカットを残す
	[string]$CSVPath,			# CSV 出力 Path
	[string]$LogPath,			# ログ出力ディレクトリ
	[switch]$AllList,			# 全リスト出力
	[switch]$WhatIf				# テスト
	)

# 全ファイルデータ名
$GC_AllFileName = "AllData"

# 重複ファイルデータ名
$GC_DuplicateFileName = "DuplicateData"

# ログの出力先
if( $LogPath -eq [string]$null ){
	$GC_LogPath = Convert-Path .
}
else{
	$GC_LogPath = $LogPath
}

# ログファイル名
$GC_LogName = "FindDuplicateFile"

##########################################################################
# ログ出力
##########################################################################
function Log(
			$LogString
			){

	$Now = Get-Date

	# Log 出力文字列に時刻を付加(YYYY/MM/DD HH:MM:SS.MMM $LogString)
	$Log = $Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " "
	$Log += $LogString

	# ログファイル名が設定されていなかったらデフォルトのログファイル名をつける
	if( $GC_LogName -eq $null ){
		$GC_LogName = "LOG"
	}

	# ログファイル名(XXXX_YYYY-MM-DD.log)
	$LogFile = $GC_LogName + "_" +$Now.ToString("yyyy-MM-dd") + ".log"

	# ログフォルダーがなかったら作成
	if( -not (Test-Path $GC_LogPath) ) {
		New-Item $GC_LogPath -Type Directory
	}

	# ログファイル名
	$LogFileName = Join-Path $GC_LogPath $LogFile

	# ログ出力
	Write-Output $Log | Out-File -FilePath $LogFileName -Encoding Default -append

	# echo
	[System.Console]::WriteLine($Log)
}

###################################################
# 指定ディレクトリ以下のファイル一覧取得
###################################################
function GetFileNames( [string]$Path ){
	if( $Recurse ){
		[array]$FileAndDirs = Get-ChildItem $Path -Recurse
	}
	else{
		[array]$FileAndDirs = Get-ChildItem $Path
	}
	[array]$TempFiles = $FileAndDirs | ? Attributes -notmatch "Directory"

	# ショートカットは対象外にする
	[array]$Files = $TempFiles | ? Name -notlike "*.lnk"

	return $Files
}

###################################################
# 指定ファイルパターン抽出
###################################################
filter SelectFiles{

	foreach($SelectPattern in $Pattern){
		if( $_.Name -like $SelectPattern ){
			return $_
		}
	}
}

###################################################
# 比較ファイル名作成
###################################################
function GetCompareFileName([string]$FileName){
	$TempFileName = $FileName -replace " - コピー", ""
	$TempFileName = $TempFileName -replace " - Copy", ""
	$TempFileName = $TempFileName -replace " \([0-9]+\)\.", "."

	return $TempFileName
}



###################################################
# 必要データ取得
###################################################
filter GetFileData{

	# Full Path
	$OriginalFileFullPath = $_.FullName

	Log "[INFO] Get detail infomation : $OriginalFileFullPath"

	if( Test-Path $OriginalFileFullPath ){
		$FileData = New-Object PSObject | Select-Object `
													CompareFileName,		# 比較ファイル名
													Hash,					# ハッシュ値
													FullPath,				# Full Path
													OriginalFileName,		# オリジナルファイル名
													LastUpdate, 			# 最終更新日付
													Size,					# サイズ(KB)
													OriginalFileNameLength,	# オリジナルファイル名の長さ
													FullPathLength,			# フルパスの長さ
													Operation,				# 操作
													BackupdFileName			# バックアップ先ファイル名

		# 比較ファイル名
		$FileData.CompareFileName = GetCompareFileName $_.Name

		# ハッシュ値
		$FileData.Hash = (Get-FileHash -Algorithm SHA256 -Path $_.FullName).Hash

		# Full Path
		$FileData.FullPath = $OriginalFileFullPath

		# オリジナルファイル名
		$FileData.OriginalFileName = $_.Name

		# 最終更新日付
		$FileData.LastUpdate = $_.LastWriteTime

		# サイズ(KB)
		$FileData.Size = [int]($_.Length / 1KB)

		# バックアップ先ファイル名
		$FileData.BackupdFileName = [string]$null

		# オリジナルファイル名の長さ
		$FileData.OriginalFileNameLength = $_.Name.Length

		# フルパスの長さ
		$FileData.FullPathLength = $OriginalFileFullPath.Length

		return $FileData
	}
	else {
		Log "[ERROR] $OriginalFileFullPath is not found !!"
	}
}

###################################################
# 重複ファイル検出
###################################################
filter KeyBreak{

	BEGIN{
		# 初期値設定
		$InitFlag = $true
		$FirstNameFlag = $true
		$FirstHashFlag = $true

		$NewKeyCompareFileName = [string]$null
		$NewKeyHash = [string]$null
		$NewRec = $null
	}

	PROCESS{
		### 通常処理

		# 初期処理
		# if( $InitFlag -eq $true ){
		#	$NewKeyCompareFileName = $File.CompareFileName
		#	$NewKeyHash = $File.Hash
		#	$NewRec = $File
		#	$InitFlag = $false
		# }

		# 比較キーとデータセット
		$OldKeyCompareFileName = $NewKeyCompareFileName
		$OldKeyHash = $NewKeyHash
		$OldRec = $NewRec

		$NewKeyCompareFileName = $_.CompareFileName
		$NewKeyHash = $_.Hash
		$NewRec = $_

		# ファイル名重複
		if( $OldKeyCompareFileName -eq $NewKeyCompareFileName ){
			$TmpRec = $OldRec
			if( $FirstNameFlag -eq $true ){
				$FirstNameFlag = $false
			}
			else{
				$TmpRec.CompareFileName = [string]$null
			}

			# ハッシュも重複
			if( $OldKeyHash -eq $NewKeyHash ){
				if( $FirstHashFlag -eq $true ){
					$FirstHashFlag = $false
				}
				else{
					$TmpRec.Hash = [string]$nul
				}
			}
			# ハッシュ ブレーク
			else{
				if( $FirstHashFlag -eq $false ){
					$TmpRec.Hash = [string]$nul
				}
				$FirstHashFlag = $true
			}

			# 重複データ
			return $TmpRec
		}
		# キーブレーク
		else{
			# キーブレークした後は重複データを出力する
			if( $FirstNameFlag -eq $false ){
				$TmpRec = $OldRec

				# ファイル名重複
				$TmpRec.CompareFileName = [string]$null

				# ハッシュも重複
				if( $FirstHashFlag -eq $false ){
						$TmpRec.Hash = [string]$nul

				}
				# フラグクリア
				$FirstNameFlag = $true
				$FirstHashFlag = $true

				# 重複データ
				return $TmpRec
			}
		}
	}

	END{
		# 残ったデーターを出力
		# キーブレークした後は重複データを出力する
		if( $FirstNameFlag -eq $false ){

			$TmpRec = $NewRec

			# ファイル名重複
			$TmpRec.CompareFileName = [string]$null

			# ハッシュも重複
			if( $FirstHashFlag -eq $false ){
					$TmpRec.Hash = [string]$nul
			}

			# 重複データ
			return $TmpRec
		}
	}
}


###################################################
# 全ファイル取得
###################################################
filter GetAllFiles{
	Log "[INFO] Get files: $_"
	if( -not (Test-Path $_ )){
		Log "[ERROR] $_ is not found !!"
	}

	return GetFileNames $_
}

###################################################
# Sort
###################################################
function DataSort($TergetFilesData){
	[array]$SortFilesData = $TergetFilesData | Sort-Object -Property `
													CompareFileName,
													Hash,
													FullPathLength,
													OriginalFileNameLength,
													OriginalFileName
	return $SortFilesData
}

###################################################
# 全データ出力
###################################################
function OutputAllData([array]$DuplicateFiles, $Now){
	$OutputFile = Join-Path $CSVPath ($GC_AllFileName + "_" +$Now.ToString("yyyy-MM-dd_HH-mm") + ".csv")

	Log "[INFO] Output all file list : $OutputFile"

	if( -not(Test-Path $CSVPath)){
		mdkdir $CSVPath
	}
	$SortFilesData | Select-Object `
						CompareFileName,
						Hash,
						FullPath,
						OriginalFileName,
						LastUpdate | Export-Csv -Path $OutputFile -Encoding Default
}


###################################################
# 重複データ出力
###################################################
function OutputDuplicateData([array]$SortFilesData, $Now){

	$OutputFile = Join-Path $CSVPath ($GC_DuplicateFileName + "_" +$Now.ToString("yyyy-MM-dd_HH-mm") + ".csv")

	Log "[INFO] Output duplicate file list : $OutputFile"

	if( -not(Test-Path $CSVPath)){
		mdkdir $CSVPath
	}
	$DuplicateFiles | Select-Object `
							CompareFileName,
							Hash,
							FullPath,
							OriginalFileName,
							LastUpdate,
							Operation,
							BackupdFileName | Export-Csv -Path $OutputFile -Encoding Default
}


###################################################
# ファイル操作
###################################################
function FileOperation( [array]$DuplicateFiles ){

	$DuplicateFileCount = $DuplicateFiles.Count
	$OperationCount = 0

	# ショートカット作成用のオリジナルフルパス
	$OriginalFilePath = [string]$null

	for( $i = 0; $i -lt $DuplicateFileCount; $i++ ){
		# ファイル名重複
		if( $DuplicateFiles[$i].CompareFileName -eq [string]$null ){

			# 重複したファイル名
			$DuplicateFile = $DuplicateFiles[$i].FullPath

			# ファイル重複
			if( $DuplicateFiles[$i].Hash -eq [string]$null ){

				# オペレーション : Move
				if( $Move -ne [string]$null ){

					if( -not (Test-Path $Move)){
						md $Move
					}
					$DuplicateFiles[$i].Operation = "Move"

					# Default 移動先ファイル名
					$MoveDdestinationFileFullName = Join-Path $Move $DuplicateFiles[$i].OriginalFileName

					# Default 移動元ファイル名
					$MoveSourceFileFullName = $DuplicateFile

					# 加工用移動先ファイル名
					$CompareFileName = GetCompareFileName $DuplicateFiles[$i].OriginalFileName
					$TempBuffer = $CompareFileName.Split( "." )
					$Ext = "." + $TempBuffer[$TempBuffer.Count -1]
					$Body = $CompareFileName.Replace($Ext, "")
					$SourceDirectory = Split-Path $DuplicateFile -Parent

					# 移動先重複回避
					$Index = 0
					while($true){
						# 重複なし
						if( -not (Test-Path $MoveDdestinationFileFullName)){
							# ファイル移動
							if( -not $WhatIf ){
								Move-Item $MoveSourceFileFullName $MoveDdestinationFileFullName
								# ショートカットを残す
								if( $CreateShortcut ){
									CreateShortcut $MoveSourceFileFullName $OriginalFilePath
								}
							}
							Log "[INFO] File duplicate (Move) : $DuplicateFile"
							break
						}


						# 移動先重複なのでファイル名をインクリメントする
						$Index++
						$MovedFileName = $Body + " (" + $Index + ")" + $Ext

						# 移動先ファイル名 をFull Path にする
						$MoveDdestinationFileFullName = Join-Path $Move $MovedFileName
					}

					# 移動先ファイル名
					$DuplicateFiles[$i].BackupdFileName = $MoveDdestinationFileFullName

					$OperationCount++
				}
				# オペレーション : Remove
				elseif( $Remove ){
					$DuplicateFiles[$i].Operation = "Remove"

					if( -not $WhatIf ){
						# 削除
						Remove-Item $DuplicateFile
						# ショートカットを残す
						if( $CreateShortcut ){
							CreateShortcut $DuplicateFile $OriginalFilePath
						}
					}
					Log "[INFO] File duplicate (Remove) : $DuplicateFile"

					$OperationCount++
				}
			}
			# ファイル名のみ重複はオリジナルファイル判定
			else{
				# Log "[INFO] Name duplicate (NOP) : $DuplicateFile"
				$DuplicateFiles[$i].Operation = "Original"
				$OriginalFilePath = $DuplicateFiles[$i].FullPath
			}
		}
		# オリジナルファイル判定
		else{
			$DuplicateFiles[$i].Operation = "Original"
			$OriginalFilePath = $DuplicateFiles[$i].FullPath
		}
	}

	return $OperationCount
}

###################################################
# 拡張子から関連付けられたプログラムを得る
# 書いたけど、使わないかも....
# 不要だとわかったら削除
###################################################
function GetExt2App([string]$Ext){
	if($Ext[0] -ne "."){
		$Ext = "." + $Ext
	}

	# 関連付けられた内部情報
	$TempBuffer = cmd /c assoc $Ext
	if( $LastExitCode -ne 0 ){
		return $null
	}

	if( $TempBuffer -like "*file"){
		return $null
	}

	$TempInfo = $TempBuffer.Split("=")
	if( $TempInfo.Count -ne 2 ){
		return $null
	}

	$InternalInfo = $TempInfo[1]

	# 関連付けられたプログラム
	$TempBuffer = cmd /c ftype $InternalInfo
	$TempInfo = $TempBuffer.Split("=")
	if( $TempInfo.Count -ne 2 ){
		return $null
	}
	# プログラムとオプション
	$TempBuffer = $TempInfo[1]

	# プログラム名だけにする
	if( $TempBuffer -match	"^`"(?<ExeInfo>.+?)`"" ){
		$ExeInfo = $Matches.ExeInfo
	}
	return $ExeInfo
}

###################################################
# ファイルのショートカットを作る
###################################################
function CreateShortcut([string]$ShortcutPath, [string]$LinkPath ){
	if($LinkPath -ne [string]$null ){

		$ShortcutName = $ShortcutPath + ".lnk"

		# ショートカットを作る
		$WsShell = New-Object -ComObject WScript.Shell
		$Shortcut = $WsShell.CreateShortcut($ShortcutName)
		$Shortcut.TargetPath = $LinkPath
		$Shortcut.Save()
	}
}

###################################################
# main
###################################################
Log "[INFO] ============== START =============="

if( $CSVPath -eq [string]$null){
	$CSVPath = Convert-Path .
}

# 指定ディレクトリ以下のファイル一覧取得
if( $Path.Count -eq 0 ){
	[array]$Path = Convert-Path .
}
[array]$AllFiles = $Path | GetAllFiles

$AllFilesCount = $AllFiles.Count
Log "[INFO] All files count : $AllFilesCount"

# ファイルパターンで対象ファイルを絞る
Log "[INFO] Select terget file."

if( $Pattern.Count -ne 0 ){
	$Patterns = ""
	$Pattern | %{ $Patterns += $_ + " " }
	Log "[INFO] Pattern select: $Patterns"
	[array]$TergetFiles =  $AllFiles | SelectFiles
}
else{
	Log "[INFO] All file."
	[array]$TergetFiles = $AllFiles
}

# 対象ファイルに Hash などの必要情報を追加
Log "[INFO] Get detail infomation."
$TergetFilesData = $TergetFiles | GetFileData

# 対象ファイル数
$TergetFilesDataCount = $TergetFilesData.Count

# 対象0件なら処理しない
if( $TergetFilesDataCount -eq 0 ){
	Log "[INFO] Terget files is zero."
}
else{
	# 対象ファイル数表示
	Log "[INFO] Terget files count : $TergetFilesDataCount"

	# Data Sort
	Log "[INFO] Data sort."
	[array]$SortFilesData = DataSort $TergetFilesData

	# 出力ファイル用処理時間
	$Now = Get-Date

	# 全ファイルリスト出力
	if( $AllList -eq $true ){
		Log "[INFO] Output all data"
		OutputAllData $SortFilesData $Now
	}

	# 重複ファイル検出
	Log "[INFO] Get duplicate files."
	[array]$DuplicateFiles = $SortFilesData | KeyBreak

	# 重複ファイル数
	$DuplicateFileCount = $DuplicateFiles.Count

	# 重複 0 件なら処理しない
	if($DuplicateFileCount -eq 0){
		Log "[INFO] Duplicate file is zero."
	}
	else{
		# 重複ファイル数表示
		Log "[INFO] Duplicate file count : $DuplicateFileCount"

		# 重複ファイル操作
		Log "[INFO] File operation"
		$Counter = FileOperation $DuplicateFiles

		if( $Counter -ne 0 ){
			Log "[INFO] File deduplication count : $Counter"
		}

		# 重複データ出力
		OutputDuplicateData $DuplicateFiles $Now
	}
}

Log "[INFO] ============== END =============="
