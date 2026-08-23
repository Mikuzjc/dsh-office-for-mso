# dsh-office-for-mso

> [中文](README.md) · [English](README.en.md) · 日本語

**DSH ↔ Office ブリッジ実行エンジン（v1.3）**：DSH セッションからコマンドを発行し、Office アドイン経由で**現在開いている** Word / Excel / PowerPoint ドキュメントを操作します——読み取り、書き込み、書式設定、構造操作、グラフ・数式・コメントなど、Microsoft Copilot for Office の一般的なワークフローに迫ります。

```
あなた ──DSHセッション──▶ AI(agent) ──POST──▶ ブリッジサービス localhost:3000
                                                    │ コマンドキュー（host ルーティング）
                           Office アドイン（バックグラウンドで1秒毎ポーリング）
                                                    │ Office.js 実行（33 アクション）
                                                    ▼
                              結果返却 ──▶ AI が取得 ──▶ あなたへ報告
```

- **競合なし**：アドインは Office プロセス内で動作し、メモリ上のドキュメントを操作。あなたの編集と Office が直列処理します（ファイルロックなし、「最後に保存した人が勝ち」なし）
- **タスクペイン操作不要**：ペインはステータス表示のみ。すべてのコマンドは DSH から送信
- **ホットリロード**：全アクションの実装は `actions.js` に集約。編集 → ブリッジ再起動 → ペインが自動再読込、**ペインを開き直す必要なし**

---

## 1. セットアップ（初回のみ）

### 1.1 リポジトリをクローンしてブリッジサービスを起動

```powershell
git clone https://github.com/Mikuzjc/dsh-office-for-mso.git
cd dsh-office-for-mso
node server.js   # または：powershell -ExecutionPolicy Bypass -File start.ps1
```

`listening on http://127.0.0.1:3000` が表示されれば OK。
**Windows では起動時にアドインが自動登録（WEF レジストリ）されます**——初回は**Office ドキュメントを閉じて開き直す**とペインが表示されます。macOS では Office メニューから手動で読み込んでください（1.2/1.3 参照）。
（以降のパス例はすべてご自身の実際のプロジェクトディレクトリを指します。）

### 1.2 アドインを Office にサイドロード

> **プラットフォーム**：中核（server.js + アドイン）は Node.js のみ必要で Windows / macOS で動作（macOS の Office デスクトップもアドインのサイドロードに対応）。`install.ps1`/タスクスケジューラは Windows 専用の**任意**の自動ホスティング。macOS では手動で `node server.js` を実行。

> **Windows ユーザー：手動サイドロード不要**——`node server.js`（または `npm run setup`）の起動時にアドインが**自動登録**されます（WEF レジストリ、通常権限で OK）。初回は**Office ドキュメントを閉じて開き直す**とペインが表示されます。
> 手動管理が必要な場合のみ：
> ```powershell
> cd <ご自身のプロジェクトディレクトリ>
> powershell -ExecutionPolicy Bypass -File sideload.ps1   # 手動登録（-Remove で削除）
> ```

**macOS ユーザー**：自動登録はありません。Office メニューから手動で読み込んでください（下記の開発者アドインの手順）：

1. 開発者タブを有効化：**ファイル → オプション → リボンのカスタマイズ → メインタブで「開発者」にチェック** → OK
2. 任意の Word / Excel / PowerPoint ドキュメントを開く
3. **開発者タブ → アドイン**（または **挿入 → アドイン**）→「Office アドイン」ダイアログを開く
4. ダイアログ左下の「管理」ドロップダウンで **開発者アドイン** を選択
5. **+（追加）** をクリック → `<ご自身のプロジェクトディレクトリ>\manifest.xml` を選択
6. サイドバーに「DSH Office 実行エンジン」ペインが表示され、**接続済み: DSH コマンド待機中** となれば成功

その後は**ドキュメントを開いたまま**にしてください。ペインは最小化・隅に移動して構いません。

> 旧バージョンの Office で「アドインのアップロード」エントリがある場合は直接使用可能。
> 別のマシンへ：上記2ステップ（ブリッジサービス + manifest アップロード）を繰り返すだけです。

### 1.3 初回利用

インストール後、DSH がドキュメントを操作できるようになる前に**アドインを一度手動で開く**必要があります：

1. Word / Excel / PowerPoint ドキュメントを開く
2. **ホーム（または開発者）タブ → アドイン → 開発者アドイン → 「DSH Office 実行エンジン」**
3. ペインが表示され（**接続済み: DSH コマンド待機中** と表示）、DSH セッションからコマンドを発行可能になります
4. ペインは小さくリサイズ・隅に移動可能（ステータス表示のみ）

その後はドキュメントを開いたままにしてください。ペインを閉じた場合は手順2を繰り返して再表示します。

## 2. アーキテクチャ

| ファイル | 役割 |
|---|---|
| `server.js` | ブリッジサービス：コマンドキュー、host ルーティング、ハートビート、静的ファイル、能力検出 `/office/capabilities`、actions バージョン `/office/actions-version` |
| `taskpane.js` | アドインシェル：ポーリング/ハートビート/ディスパッチ/ホットリロード読み込み（**変更は最小限に**、変更時はペイン再オープンが必要） |
| `actions.js` | 全アクション実装 + レジストリ（**ホットリロード対象**：編集時はペイン再オープン不要） |
| `pako.min.js` | ローカル zip 解凍ライブラリ（PPT OOXML 読み取り用、オフライン可） |
| `manifest.xml` | アドインマニフェスト（権限 ReadWriteDocument、最高レベル） |

**マルチドキュメントモデル**：Word / Excel / PowerPoint はそれぞれ独立したアドインインスタンスを実行。コマンドは `host`（Word/Excel/PowerPoint）で正確にルーティングされ、`GET /office/status` がオンラインドキュメント一覧（`hosts` フィールド）とペイン起動記録（`hellos`）を返します。

**ホットリロードの仕組み**：ペインはポーリングのたびに `/office/actions-version` を GET し、`actions.js` の mtime と比較。変更があればスクリプトを動的に再読み込みします。**`actions.js` を編集 → サーバー再起動 → 自動反映**。

## 3. 能力マトリクス（33 アクション）

> `destructive=true` の操作は `args.dryRun` で影響をプレビュー可能（replace_all / remove_empty_paragraphs / delete_sheet で実装済み、他は AI 層で「先読み後書き」）。W=Word、E=Excel、P=PowerPoint。

### 共通
| action | プラットフォーム | 説明 |
|---|---|---|
| `read_selection` | W/E/P | 現在の選択テキストを読み取り；`withStyles=true` でスタイルも返却 |
| `write_selection` | W/E/P | 選択範囲を `{text}` で置換 |
| `read_document` | W/E/P | Word 全文；Excel はワークシートの使用範囲（`sheet` で指定可、5000セル上限）；PPT は全ファイルをスライド毎に |
| `read_styles` | W/E | 選択スタイル：Word（フォント/サイズ/太字/斜体/色/下線/ハイライト）；Excel（セル毎、最大10×10） |
| `replace_all` | W/E | 文書全体の検索置換 `{search, replace, dryRun?}` |
| `append_text` | W | 文末に段落を追加 `{text}` |

### Word グループ
| action | 説明 |
|---|---|
| `read_tables` | 全テーブルを構造化読み取り（セル毎、getRange().text を \t/\r で分割） |
| `set_font` | 文書全体（表内段落含む）のフォント設定 `{font}` |
| `remove_empty_paragraphs` | 空段落を削除（**画像段落と文書末尾段落はスキップ**、`dryRun` プレビュー可） |
| `insert_paragraph` | 段落挿入 `{text, style?, location?}`（スタイル：見出し1-3/本文/引用/強調） |
| `insert_table` | テーブル挿入 `{rows}` ⚠️ **この環境ではテーブル挿入 API がすべて利用不可（能力制限参照）** |
| `insert_image` | 選択位置に画像挿入 `{base64, width?, height?}` |
| `apply_style` | 組み込みスタイル適用 `{style, scope: selection\|all}` |
| `format_selection` | 選択範囲の書式設定 `{font, size, bold, italic, color, highlight}` |
| `set_paragraph_format` | 段落書式 `{alignment, indent, lineSpacing, listType}` ⚠️ **この環境では paragraphFormat 利用不可** |
| `search` | 検索 `{query, matchCase?, wildcard?}`、ヒット一覧を返却 |
| `add_comment` | 選択範囲にコメント追加 ⚠️ **この環境では Word コメント API 利用不可** |
| `read_comments` | コメント一覧 ⚠️ 同上 |
| `read_properties` | ドキュメントプロパティ（タイトル/作成者/文字数など） |

### Excel グループ
| action | 説明 |
|---|---|
| `list_sheets` | ワークシート一覧（名前/位置/表示状態） |
| `read_range` | 範囲読み取り `{address: "Sheet1!A1:B10", limit?}`（値/数式/表示形式） |
| `write_range` | 一括書き込み `{address, values?/formulas?}`（2次元配列、>5000セルは自動分割） |
| `format_range` | 範囲書式設定 `{address, font, size, bold, fill, numberFormat, autoFit, tableStyle}` |
| `insert_chart` | データ→グラフ `{type: Column/Line/Bar/Pie/Area/Scatter/…, dataRange, title?}` |
| `add_sheet` / `rename_sheet` / `delete_sheet` | ワークシート管理（delete は dryRun 対応） |
| `apply_sort` | 並べ替え `{address, fields: [{column, ascending}]}` |
| `apply_filter` | オートフィルター `{address, columns?}` |
| `evaluate_formula` | 数式評価 `{formula: "SUM(A1:A10)"}`（ホワイトリスト SUM/AVERAGE/COUNT/MAX/MIN/PRODUCT、workbook.functions で型付き評価） |
| `add_comment` / `read_comments` | セルコメント（`cell` はシート名プレフィックス付き可、自動除去） |
| `read_properties` | ワークブックプロパティ |

### PPT グループ
| action | 説明 |
|---|---|
| `read_slides` | 現在選択中のスライド一覧（SlideRange：id + title） |
| `ppt_read_notes` | 全ファイルの発表者ノート（OOXML notesSlides 解析 + rels マッピング） |
| `read_document` | 全ファイルをスライド毎にテキスト化（共通） |

### 環境診断
| action | 説明 |
|---|---|
| `get_environment` | ホストバージョン/プラットフォーム/requirementSets 対応状況 + Word オブジェクトモデルの深層プローブ（能力制限の特定用） |

## 4. 能力制限（実測結果、2026-08）

**帰属の明確化**：「利用不可」は2種類に分類されます——
- **プラットフォーム不可能**（Office.js 仕様に API が存在せず、どのバージョン/マシンでも不可）：PPT の OOXML 書き込み、Word のページ設定/目次/クリップボード移動、ペインのプログラム的更新
- **このマシンのランタイム欠落**（requirementSets は WordApi 1.8 / ImageCoercion 1.1 を宣言、しかしランタイムのオブジェクトプロパティが実際に欠落；node.js の問題でも CDN キャッシュの問題でもない——検証済み）：Word コメント（`body.comments` プロパティ不存在）、`paragraphFormat`（プロパティ不存在）、テーブル挿入（`insertTable`/`insertOoxml` は存在するが呼び出し失敗）。**他のマシン/新しい Office では動作する可能性が高い**。本プロジェクトは `requirement` エラーコードを返し誠実に劣化動作します

| プラットフォーム | 利用可 | 利用不可とカテゴリ |
|---|---|---|
| Word | 段落/テキスト挿入・置換、全文/テーブル読み取り、フォント、スタイル、選択書式、検索、ドキュメントプロパティ、空段落クリーンアップ | テーブル挿入 / paragraphFormat / コメント（**ランタイム欠落**）；ページ設定/目次/クリップボード（**プラットフォーム**） |
| Excel | ワークシート管理、範囲読み書き（一括）、書式、グラフ、並べ替え、フィルター、数式評価、コメント、プロパティ | `workbook.getRange`/`calculate` 不存在（代替 API で回避）；Range.autoFilter はワークシートレベル API が必要 |
| PPT | 全ファイルのテキスト/ノート読み取り、SlideRange | OOXML 書き込み、スライド追加/レイアウト（**プラットフォーム**） |

**設計原則**：未対応 API は `code: requirement | unsupported | execution` を返し、AI 層が劣化動作または誠実に報告——成功を偽装しません。

## 5. セーフティガードレール

- **破壊的操作の dryRun**：replace_all / remove_empty_paragraphs / delete_sheet は `dryRun` で影響プレビュー。AI 層は実行前にプレビュー
- **画像段落保護**：空段落削除時、inlinePictures を含む段落はスキップ（過去にフローチャートを誤削除した事故を修正済み）
- **文末段落保護**：Word の最後の段落（段落記号）は削除不可
- **パフォーマンス保護**：Excel 一括書き込みは分割（≤5000セル/回）、`getUsedRange(true)` で全列書式の爆発を回避、大規模読み取りは切り詰め

## 6. サービスホスティング（本番）

**推奨：Windows タスクスケジューラ「DSH Office Bridge」（ログオン時自動起動、サイレント）**
- **ワンクリックインストール**：`powershell -ExecutionPolicy Bypass -File install.ps1`（カレントディレクトリを自動で使用して登録、マシン非依存でパス編集不要）
- 手動登録：トリガー=ユーザーログオン時；設定=時間制限なし（常駐）、StartWhenAvailable；起動コマンド=`powershell -NoProfile -WindowStyle Hidden -Command "& '<node完全パス>' '<プロジェクトディレクトリ>\server.js'"`（サイレント、ウィンドウなし）
- 手動管理：
  - 起動：`Start-ScheduledTask -TaskName 'DSH Office Bridge'`
  - 停止：ポートでプロセス特定 `netstat -ano | findstr :3000` → `Stop-Process -Id <pid>`
  - 再起動（コード変更後）：プロセス停止 → `Start-ScheduledTask -TaskName 'DSH Office Bridge'`

**開発中**：`powershell -ExecutionPolicy Bypass -File start.ps1`（フォアグラウンド）；または `npm start`。

**自己チェック**：`powershell -ExecutionPolicy Bypass -File smoke-test.ps1`（サービス/エンドポイント/オンラインドキュメントを確認）。

## 7. トラブルシューティング

| 症状 | 対処 |
|---|---|
| ペインに「ブリッジサービス未接続」 | `node server.js` が実行中か確認（`/office/status` が応答するか） |
| DSH コマンドがタイムアウト | アドインがオフライン：ドキュメントが開いているか + ペインが「接続済み」か確認 |
| ペインが表示されない | manifest を再アップロード；Office が管理者権限でないか確認（localhost 例外は通常権限が必要） |
| 変更が反映されない | カーソル/選択位置を確認；ペインの実行ログを確認 |
| actions.js の変更が反映されない | ブリッジを再起動（ペインはホットリロード、再オープン不要） |
| taskpane.js の変更が反映されない | ペインを手動で開き直す（シェルコードはプログラム的に更新不可、Office デスクトップの制限） |

## 8. 能力検出（AI 側）

- `GET /office/capabilities` → アクションレジストリ（名前/プラットフォーム/破壊的か/引数説明）
- `GET /office/status` → オンラインドキュメント（hosts）+ ペイン起動記録（hellos）
- `GET /office/actions-version` → actions.js バージョン（ホットリロード比較用）

---

*Microsoft 公式製品ではありません。`DSH` はコミュニティ AI ハーネスエコシステムであり、本プロジェクトは DSH と Microsoft Office を結ぶ独立したブリッジです。*
*Vibe-coded：人間と AI のペア開発（AI 支援コーディング）により作成され、全機能を実ドキュメントで検証済み。*
