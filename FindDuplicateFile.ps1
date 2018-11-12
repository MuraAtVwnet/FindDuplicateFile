###################################################
# 重複ファイル検出
###################################################
Param(
	[array][string]$Paths,		# 探査する Path
	[array][string]$Patterns,	# ファイルパターン
	[string]$CSVPath,			# CSV 出力 Path
	[switch]$Remove,			# 削除実行
	[string]$BackupDirectory,	# バックアップフォルダ
	[string]$LogPath,			# ログ出力ディレクトリ
	[switch]$AllList			# 全リスト出力
	)

# 全ファイルデータ名
$GC_AllFileName = "AllData"

# 重複ファイルデータ名
$GC_DuplicateFileName = "DuplicateData"

# ログの出力先
if( $LogPath -eq [string]$null ){
	$GC_LogPath = $PSScriptRoot
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
	Write-Host $Message $Log
}

###################################################
# 指定ディレクトリ以下のファイル一覧取得
###################################################
function GetFileNames( [string]$Path ){
	[array]$FileAndDirs = Get-ChildItem $Path -Recurse
	[array]$Files = $FileAndDirs | ? Attributes -notmatch "Directory"
	return $Files
}

###################################################
# 指定ファイルパターン抽出
###################################################
filter SelectFiles{

	foreach($SelectPattern in $Patterns){
		if( $_.Name -like $SelectPattern ){
			return $_
		}
	}
}


###################################################
# 必要データ取得
###################################################
filter GetFileData{

	# Full Path
	$OriginalFileFullPath = $_.FullName

	Log "[INFO] Get hash data : $OriginalFileFullPath"

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
		$TempFileName = $_.Name -replace " - コピー", ""
		$TempFileName = $TempFileName -replace " - Copy", ""
		# $TempFileName = $TempFileName -replace "_[0-9]+\.", "."
		$FileData.CompareFileName = $TempFileName -replace " \([0-9]+\)\.", "."

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
				# 重複データ
				return $TmpRec
			}
			# フラグクリア
			$FirstNameFlag = $true
			$FirstHashFlag = $true
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
# main
###################################################
Log "[INFO] ============== START =============="

if( $CSVPath -eq [string]$null){
	$CSVPath =$PSScriptRoot
}

# 指定ディレクトリ以下のファイル一覧取得
if( $Paths.Count -eq 0 ){
	[array]$Paths = Convert-Path .
}
[array]$AllFiles = $Paths | GetAllFiles

$AllFilesCount = $AllFiles.Count
Log "[INFO] All files count : $AllFilesCount"

# ファイルパターンで対象ファイル抽出
Log "[INFO] Select terget file."

if( $Patterns.Count -ne 0 ){
	Log "[INFO] Pattern select."
	[array]$TergetFiles =  $AllFiles | SelectFiles
}
else{
	Log "[INFO] All file."
	[array]$TergetFiles = $AllFiles
}

# 対象ファイルに Hash などの必要データを追加
Log "[INFO] Add hash data"
$TergetFilesData = $TergetFiles | GetFileData

$TergetFilesDataCount = $TergetFilesData.Count
Log "[INFO] Terget files count : $TergetFilesDataCount"

# Sort
Log "[INFO] Sort file datas"
[array]$SortFilesData = $TergetFilesData | Sort-Object -Property `
												CompareFileName,
												Hash,
												FullPathLength,
												OriginalFileNameLength,
												OriginalFileName

$Now = Get-Date

# 全ファイルリスト出力
if( $AllList -eq $true ){
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

# 重複ファイル検出
Log "[INFO] Get duplicate files."
[array]$DuplicateFiles = $SortFilesData | KeyBreak

# ファイル処理
if( $BackupDirectory -ne [string]$null ){
	# 移動
	# オペレーション : Move
	# 移動先ファイル名
}
elseif( $Remove ){
	# 削除
	# オペレーション : delete
}

# 重複データ出力(テスト用/仕上げるときはオペレーションで必要フィールド追加するので、全オブジェクト出力)
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

Log "[INFO] ============== END =============="
