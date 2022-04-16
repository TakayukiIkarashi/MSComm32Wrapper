# MSComm32Wrapper
## 概要
QRコードリーダーを使うためのCOMコンポーネントです。  
VB6で開発したDLLのため、レジストリに登録して利用します。
## 使い方
MSComm32Wrapper.dllをレジストリに登録して利用します。  
MSComm32Wrapper.dllを任意のフォルダに配置し、以下のコマンドをコマンドプロンプトより実行します。

    regsvr32 [MSComm32Wrapper.dllのフルパス]

MSComm32Wrapper.dllをレジストリから削除する場合は、以下のコマンドをコマンドプロンプトより実行します。

    regsvr32 /u [MSComm32Wrapper.dllのフルパス]
