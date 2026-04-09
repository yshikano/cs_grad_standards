# cs_grad_standards GitHub Pages runtime fix

この版は、GitHub Pages 公開時に `科目データを読み込み中…` のまま進まなくなる不具合を修正した差し替え版です。

## 主な修正

- `app.js` 内の後付けパッチ部分で、
  `const __old_x = x; function x(...) { __old_x(...); }`
  のような再定義が複数入っており、関数宣言の巻き上げの影響で自己再帰になっていました。
- 上記を、以前の関数を正しく捕まえる代入形式へ変更しました。
- `relinkDraftSourceCourses20260408()` が `draft.courseRows` を配列だと仮定していた箇所を、
  既存実装どおりオブジェクト形式でも動くように修正しました。
- `searchSideView` の初期化で `localStorage` が使えない環境でも落ちにくいように防御しました。

## 置き換え方法

既存の GitHub リポジトリ `yshikano/cs_grad_standards` では、最低限次の 1 ファイル差し替えで動きます。

- `app.js`

確実に反映したい場合は、このフォルダの中身でリポジトリ公開用ファイルを丸ごと置き換えてください。

## 同梱ファイル

- `index.html`
- `app.js`  ← 修正版
- `styles.css`
- `data/courses.sample.json`
- `data/kdb-grad.json`
- `.nojekyll`

## 確認済みポイント

- 科目データが読み込まれ、検索件数が表示されること
- `data/kdb-grad.json` を使って全大学院科目カタログが統合されること
- GitHub Pages 配置想定の相対パス `./data/...` で動くこと
