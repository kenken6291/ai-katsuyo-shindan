# AI活用診断（ai-katsuyo-shindan）

GitHub Pages（index.html）＋ GAS（Code.gs）＋ スプレッドシート ＋ Google Drive ＋ Gemini API

## 1. 全体フロー

```
[未ログイン]
トップ ─ 簡易診断(5問) ─ quickDiagnose ─ 結果（レベル＋最初の一歩）─ 会員登録へ

[会員登録]
register ─ 仮パスワードをメール送信（is_initial_login=TRUE）
login(仮PW) ─ mustChangePassword=true ─ パスワード設定画面（強制）─ changePassword ─ マイページ

[本診断]
マイページ ─ 個人 / 法人(小・中・大) ─ 6〜8問 ─ fullDiagnose ─ 詳細レポート
                                                          └ emailReport ─ PDFをDriveに保存＋メール送信
[パスワード忘れ]
forgotPassword ─ 新しい仮パスワード送信（旧PW無効・全セッション削除）─ 再びパスワード設定へ
```

- 通信は全て `POST`（`Content-Type: text/plain`）で `{action, token, ...}` を送信（CORSプリフライト回避）
- トークンは localStorage に保存（有効期限7日、sessionsシート＋CacheService）
- ログイン5回連続失敗で30分ロック
- パスワードは SHA-256 ＋ salt ＋ pepper（200回ストレッチ）

## 2. スプレッドシート定義

`setup()` を実行すると3シートが自動作成されます（全列テキスト書式）。

### users
| 列 | 内容 |
|---|---|
| id | ユーザーID（U＋16桁） |
| email | メールアドレス（小文字） |
| password_hash | ハッシュ化パスワード |
| salt | ソルト |
| name | 氏名 |
| company_name | 会社名（法人のみ） |
| attribute | 個人 / 法人 |
| company_size | 小規模 / 中規模 / 大規模 |
| is_initial_login | TRUE＝仮パスワード状態 |
| created_at | 作成日（ISO形式） |
| last_login_at | 最終ログイン日 |
| failed_count | 連続ログイン失敗回数 |
| locked_until | ロック解除時刻 |

### diagnostics
| 列 | 内容 |
|---|---|
| diagnosis_id | 診断ID |
| user_id | ユーザーID（未登録は guest） |
| type | 簡易 / 個人 / 法人_小規模 / 法人_中規模 / 法人_大規模 |
| answers_json | 回答データ（JSON） |
| result_json | Gemini生成結果（JSON） |
| created_at | 診断日時 |
| pdf_file_id | メール送信したPDFのDriveファイルID |

### sessions
| 列 | 内容 |
|---|---|
| token | セッショントークン |
| user_id | ユーザーID |
| created_at | 発行日時 |
| expires_at | 有効期限 |

## 3. セットアップ手順

1. Googleスプレッドシートを新規作成（名前は自由）
2. 「拡張機能 → Apps Script」を開き、`Code.gs` を全文貼り付けて保存
3. 「プロジェクトの設定 → スクリプト プロパティ」に追加
   - `GEMINI_API_KEY` … Google AI Studio で発行したキー（必須）
   - `APP_URL` … 例 `https://kenken6291.github.io/ai-katsuyo-shindan/`（メール本文用）
   - `GEMINI_MODEL` … 変更したいときのみ（既定 `gemini-2.5-flash`）
4. エディタで関数 `setup` を選んで実行 → 権限を承認（シート作成・PEPPER生成・Driveフォルダ作成）
5. 任意：`testGemini` を実行して Gemini 接続を確認
6. 「デプロイ → 新しいデプロイ → ウェブアプリ」
   - 次のユーザーとして実行：**自分**
   - アクセスできるユーザー：**全員**
7. 発行された `/exec` のURLを `config.js` の `GAS_URL` に貼り付け（最初の1回だけ）
8. GitHubリポジトリ（例 `ai-katsuyo-shindan`）に `index.html` と `config.js` をpushし、GitHub Pages を有効化
9. 任意：`cleanupSessions` に1日1回の時間主導型トリガーを設定

**GASを修正したら「デプロイを管理 → 鉛筆（編集）→ バージョン：新バージョン → デプロイ」で更新します。**
この方法ならURLは変わらないので、`config.js` の修正は不要です。
（「新しいデプロイ」を作るとURLが変わるので使わないこと）

`index.html` を差し替えても `config.js` はそのまま残せば、GAS_URLの再設定は不要です。

## 4. 補足

- メール送信は MailApp（無料アカウントは1日100通程度が上限）
- 未ログインの簡易診断は1日200回まで（`GUEST_DAILY_LIMIT` で変更）
- Geminiへは氏名・会社名・メールアドレスを送らず、属性と会社規模、回答のみ送信
- PDFは Drive の「AI活用診断レポート」フォルダに保存（共有設定は非公開のまま）
