COM_Test
========

自前のCOMに後始末用の関数をセットして、ブックを閉じる際に自動で呼ばれる様にする。

VBAはオブジェクト変数に明示的にNothingをセットしなくても自動で解放されるが、
それを逆手に取ったmoug発のテク。
Mainプロシージャを実行、後は(Excelなら)ブックを閉じれば自動的にRelease関数が呼ばれる。
