﻿●○ 重複ファイル排除スクリプト ○●

■ 概要
● FindDuplicateFile.ps1
指定フォルダにある重複ファイルの検出、削除、移動をします。
比較は、ファイル名の「- コピー」「- Copy」「(数字)」を無視し、sha256 ハッシュと合わせて比較しますので、ファイル名が同じでも内容が異なる場合は別ファイル扱いにします。

重複したファイルは CSV 出力し、削除/移動した場合はリストにその操作も記録されます。
カレントディレクトリに実行ログも出力されます。
データ出力先は default カレントディレクトリですが、指定することも可能です(ファイル名指定はできません)

● FindDuplicateHash.ps1
FindDuplicateFile.ps1 と機能的には一緒ですが、ファイル名は無視してハッシュ値が同じファイルを重複ファイルとして処理するバリエーション版です。

こちらを使う場合は、以降の説明の FindDuplicateFile.ps1 を FindDuplicateHash.ps1 と読み替えてください。


■ 使い方
・重複ファイルを削除する
FindDuplicateFile.ps1 -Remove

・重複ファイルを C:\Backup へ移動する
FindDuplicateFile.ps1 -Move  C:\Backup

■ 構文
FindDuplicateFile.ps1 [[-Path] <String[]>] [[-ExcludePath] <String[]>] [-Recurse] [[-Pattern] <String[]>] [[-ExcludePattern] <String[]>] [-Remove] [[-Move] <String>] [-CreateShortcut] [[-CSVPath] <String>] [[-LogPath] <String>] [-AllList] [-WhatIf] [<CommonParameters>]

■ 説明
重複リストのみ出力(Remove/Moveオプションを指定していない時の動作)
    重複リスト出力だけをします

探査する Path(-Path)
    探査する Path を指定します

除外する Path(-ExcludePath)
    探査しない Path を指定します

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

除外ファイルパターン指定(-ExcludePattern)
    特定拡張子のファイルを対象から除外する場合は、除外パターン(*.mov とか)を指定します
    複数指定する場合は、カンマ「,」で区切ってください

ショートカット作成(-CreateShortcut)
    ファイルを削除/移動する際にオリジナルファイルへのショートカットを残します

テスト(-WhatIf)
    実際の削除/移動はせず、動作確認だけします

■ オプション
-Path
    探索 Path
    省略時はカレントディレクトリ以下を探索します
    複数指定する場合はカンマで区切ります

-ExcludePath
    探索から除外する Path
    省略時は除外しません
    複数指定する場合はカンマで区切ります

-Recurse
    サブディレクトリも探査する

-Pattern
    対象ファイルパターン
    省略時はすべてのファイルを対象にします
    複数指定する場合はカンマで区切ります

-ExcludePattern
    対象から除外するファイルパターン
    省略時は除外しません
    複数指定する場合はカンマで区切ります

-Remove
    重複ファイルを削除します
    Remove/Move が指定されていない場合は重複リストのみを出力します

-Move
    重複ファイルの移動先
    Remove/Move が指定されていない場合は重複リストのみを出力します
    Remove/Move の両方が指定された場合は Move が優先されます

-CreateShortcut
    重複ファイルを削除/移動した時に、オリジナルファイルへのショートカットを残します
    (Windows プラットフォーム専用オプション)

-CSVPath
    CSV の出力先
    省略時はカレントディレクトリに出力します

-LogPath
    実行ログの出力先
    省略時はカレントディレクトリに出力します

-AllList
    重複確認する全ファイル リストを CSV 出力します

-WhatIf
    実際の削除/移動はせず、動作確認だけします

■ 例

PS C:\Photo> ~\FindDuplicateFile.ps1 -Recurse
カレントディレクトリ(C:\Photo)以下にある全ファイルの重複リストを出力します

PS C:\Test> .\FindDuplicateFile.ps1 -Path "C:\Photo" -Recurse
指定ディレクトリ(C:\Photo)以下にある全ファイルの重複リストを出力します

PS C:\Test> .\FindDuplicateFile.ps1 -Path "C:\Photo" -Recurse -ExcludePath "C:\Photo\Movie"
指定ディレクトリ(C:\Photo)以下にある全ファイルの重複リストを出力しますが、除外ディレクトリ(C:\Photo\Movie)は対象外にします

PS C:\Photo> ~\FindDuplicateFile.ps1 -Recurse -Pattern *.jpg, *.png
カレントディレクトリ(C:\Photo)以下にある「*.jpg」と「*.png」の重複リストを出力します

動画や ISO などサイズの大きなファイルは、ハッシュ値の取得に時間がかかるので、特定種別のファイルだけ重複排除する場合はファイルパターンを指定するがお勧めです。

PS C:\Photo> ~\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複リストを出力します

PS C:\Photo> .\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -ExcludePattern *.iso, *.mov, *.mp4
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.iso」と「*.mov」と「*.mp4」以外の重複リストを出力します

PS C:\Photo> ~\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Remove
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを削除します

PS C:\Photo> ~\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Move C:\Photo\Backup
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを「C:\Photo\Backup」へ移動します

PS C:\Photo> ~\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Move C:\Photo\Backup -WhatIf
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを「C:\Photo\Backup」へ移動した場合の動作を確認します(実際の移動はしません)

PS C:\Photo> ~\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Move C:\Photo\Backup -CSVPath C:\Photo\CSV
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを「C:\Photo\Backup」へ移動します
重複リストを「C:\Photo\CSV」に出力します

PS C:\Photo> ~\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Move C:\Photo\Backup -CSVPath C:\Photo\CSV -LogPath C:\Photo\Log
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを「C:\Photo\Backup」へ移動します
重複リストは「C:\Photo\CSV」に、実行ログは「C:\Photo\Log」に出力します

PS C:\Photo> .\FindDuplicateFile.ps1 -Path C:\Photo\2018-10, C:\Photo\2016-03 -Recurse -Pattern *.jpg, *.png -Move C:\Photo\Backup -CreateShortcut
「C:\Photo\2018-10」と「C:\Photo\2016-03」以下にある「*.jpg」と「*.png」の重複ファイルを「C:\Photo\Backup」へ移動し、代わりにショートカットを置きます
重複リストを「C:\Photo\CSV」に出力します

ショートカット(.lnk)は処理対象から外しています。

■ 動作確認環境
Windows 10 : PowerShell 5.1
Windows 10 : PowerShell 6.0.1

6.0.1 で動作確認しているので、Windows 以外でも動作すると思います。

■ ダウンロード
リポジトリを Clone するか、Invoke-WebRequest してください

Invoke-WebRequest https://raw.githubusercontent.com/MuraAtVwnet/FindDuplicateFile/master/FindDuplicateFile.ps1 -OutFile ~\FindDuplicateFile.ps1

■ リポジトリ
https://github.com/MuraAtVwnet/FindDuplicateFile
git@github.com:MuraAtVwnet/FindDuplicateFile.git

■ Web site
重複したファイル排除するスクリプト(PowerShell)
http://www.vwnet.jp/Windows/PowerShell/2018111601/FindDuplicateFile.htm

■ Feed back
Bug リポートなどありましたらfeed back をください

mura-softwaers@vwnet.jp
