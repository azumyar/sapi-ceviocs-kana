# sapi-ceviocs-kana
棒読みちゃん(SAPI5)とCeVIO CSの連携  
棒読みちゃんで話者をパラメータを変えて複数登録して遊ぶ目的のSAPIプラグインです。レジストリ一括登録バッチなどを作成してご利用ください。  
https://github.com/azumyar/sapi-ceviocs-kana/releases

インストールはリリースzip内に配置されているbatを管理権限で実行することで簡単に行えます。

このプラグインを利用するには.NET 8デスクトップランタイムが必要です。  
32bitアプリケーションから利用するにはCeVIO CS6、
64bitアプリケーションから利用するにはCeVIO CS7、
それぞれ対応したCeVIOが必要です。  
CeVIO AIには対応していません。

|  レジストリ名  |  値  |
| ---- | ---- |
|  x-cevio-cast  |  CeVIO CSキャスト  |
|  x-cevio-volume  |  音の大きさ |
|  x-cevio-speed  |  話す速さ  |
|  x-cevio-tone  |  音の高さ  |
|  x-cevio-tone-scale  |  抑揚  |
|  x-cevio-alpha  |  声質  |
|  x-cevio-components  |  感情  |
|  x-kana  |  アルファベットを仮名に変換する  |
