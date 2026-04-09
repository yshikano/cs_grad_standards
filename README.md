# cs_grad_standards startup-fix (2026-04-09)

GitHub Pages で `科目データを読み込み中…` のまま止まる症状に対する修正版です。

変更点:
- 起動直後に内蔵データで必ず画面を初期化
- `data/courses.sample.json` と `data/kdb-grad.json` は同じリポジトリ内の固定パスを優先
- `document.scripts` 依存の候補 URL 生成を廃止
- 起動失敗時に画面へエラーを表示

反映方法:
1. GitHub Pages 用リポジトリの公開フォルダを、この zip の中身で置き換える
2. `data/kdb-grad.json` と `data/courses.sample.json` が同じリポジトリにあることを確認する
3. ページをハードリロードする
