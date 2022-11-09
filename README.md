# diffxl: Excelのdiffを出力するツール

```
# 依存ライブラリをインストール
$ pipenv install
$ pipenv shell

# Excelファイルを比較
$ python diffxl.py file file
```

半分は勉強のために作ったものなので、実用性は・・・？　py-excel-textconvとかを使えるならそちらを使った方が良い
2つのファイルで入力範囲に違いがあると上手く差異を検出できないバグがある気がする
