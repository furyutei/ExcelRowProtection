ExcelRowProtection
==================

Excel で、特定の行に対するの挿入／削除を抑制するための実装例。  
[GitHubGist で公開している行挿入／削除監視用クラス](https://gist.github.com/furyutei/b31b738c0b9780a075f573eff5cef20e) の、別バージョン。  

使い方
------
1. 対象となる Excel ブック（マクロ有効ワークシート）を開き、開発→ Visual Basic で VBE を起動して、[clsRowProtection クラス モジュール](https://github.com/furyutei/ExcelRowProtection/blob/master/src/ClassModules/clsRowProtection.cls)をインポート  
1．同じく、[Sheet1 クラス モジュール](https://github.com/furyutei/ExcelRowProtection/blob/master/src/SheetModules/Sheet1.cls)をインポート（※クラス モジュールの所に入るので、シート用モジュールの雛型として使用）  
1. [サンプル Excel ブック](https://github.com/furyutei/ExcelRowProtection/blob/master/examples/RowProtectionExample.xlsm)の「RowProtectionConfig」シートを参考に設定用シート「RowProtectionConfig」を作成し、保護対象行のセル参照を記述（一枚の設定シート上に複数シート・複数行の設定を記述可能、A・B列における空欄や別シートへのセル参照数式になっていないセルは無視される）  
1. VBE にて、対象となる（行の挿入／削除保護を行いたい）シートオブジェクトに、2. でインポートされたソースコードの内容をそのままコピー＆ペースト  

制限など
--------
- 動作保証無し  
- シートもしくはテーブルの行全体を対象とする挿入／削除のみ対応（部分的な挿入／削除には未対応）  
- 上書き等に関しては保護されないため、必要に応じて別途シートの保護などをかけること  

その他
------
- ソースコードは、Excel の VBA からエクスポートしたもの（文字コード：シフトJIS・改行：CR+LF）。  

参考
----
- [端緒となったはけたさんのツイート](https://twitter.com/excelspeedup/status/1082122098791788544)  
- [VBAで行の挿入を禁止する方法 part1 提案編 - えくせるちゅんちゅん](https://kotori-chunchun.hatenablog.com/entry/2019/01/07/200341)  
- [VBAで行の挿入を禁止する方法 part2 実用編 - えくせるちゅんちゅん](https://kotori-chunchun.hatenablog.com/entry/2019/01/08/224854)  

ライセンス
----------
[The MIT License](https://github.com/furyutei/ExcelRowProtection/blob/master/LICENSE)  
