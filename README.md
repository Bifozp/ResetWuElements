# ResetWuElements

このスクリプトは、なんらかの問題によりWindows Updateができなくなった際に、
Windows Updateのクライアントを初期化する為のスクリプトです。

## このスクリプトが行うこと

- `%SystemRoot%\SoftwareDistribution` の移動・復元
- `%SystemRoot%\System32\catroot2` の移動・復元
- 上述の２フォルダが依存するサービスの停止・再開

## どのような場面で使用するか

Windows Update が、エラーコード `0xC1900204` で正常に完了しなかった場合を想定して作成されています。  
Windows Update で使用しているファイルの不整合が原因であれば、このスクリプトを利用することで復元できるかもしれません。


## 利用方法

スクリプトをダブルクリックで使用します。  
初回起動の場合、Windows Updateの初期化動作になります。

処理が成功すると、`backups` フォルダ以下にWindows Update関連ファイルを移動します。

以前にこのスクリプトを使用して初期化操作を行っている場合、バックアップの復元または削除を行います。

## 免責

このスクリプトを使用して発生したいかなる問題についても責任は負えません。

## リファレンス

- [How to fix Windows 10 Update error code 0xc1900204?](https://ugetfix.com/ask/how-to-fix-windows-10-update-error-code-0xc1900204/)
- [VBScript - VBScriptを管理者として実行する : Server World](https://www.server-world.info/query?os=Other&p=vbs&f=1)

