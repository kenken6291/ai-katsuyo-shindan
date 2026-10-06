/**
 * AI活用診断 Webアプリ - バックエンド（Google Apps Script）
 * ファイル: Code.gs
 *
 * スクリプトプロパティ（プロジェクトの設定 → スクリプト プロパティ）
 *   GEMINI_API_KEY    : 必須  Gemini APIキー
 *   GEMINI_MODEL      : 任意  既定 gemini-2.5-flash
 *   APP_URL           : 任意  GitHub PagesのURL（メール本文に記載）
 *   SPREADSHEET_ID    : 任意  スタンドアロンGASの場合のみ必須（コンテナバインドなら不要）
 *   DRIVE_FOLDER_ID   : 自動  setup() 実行時に作成
 *   PEPPER            : 自動  setup() 実行時に作成（変更すると全員ログイン不可になるので触らない）
 *   GUEST_DAILY_LIMIT : 任意  簡易診断（未ログイン）の1日あたり上限。既定 200
 *
 * ※ コードを変更したら「デプロイを管理 → 編集 → 新バージョン」で再デプロイしてください。
 */

const APP_NAME = 'AI活用診断';
const SHEETS = {
  users: ['id', 'email', 'password_hash', 'salt', 'name', 'company_name', 'attribute', 'company_size',
          'is_initial_login', 'created_at', 'last_login_at', 'failed_count', 'locked_until'],
  diagnostics: ['diagnosis_id', 'user_id', 'type', 'answers_json', 'result_json', 'created_at', 'pdf_file_id'],
  sessions: ['token', 'user_id', 'created_at', 'expires_at']
};
const SESSION_DAYS = 7;
const MAX_FAILED = 5;
const LOCK_MINUTES = 30;
const HASH_ITER = 200;
const SIZES = ['小規模', '中規模', '大規模'];
const FULL_TYPES = ['個人', '法人_小規模', '法人_中規模', '法人_大規模'];
const TYPE_LABEL = {
  '簡易': '簡易診断',
  '個人': '個人向け診断',
  '法人_小規模': '法人向け診断（小規模：1〜20名）',
  '法人_中規模': '法人向け診断（中規模：21〜100名）',
  '法人_大規模': '法人向け診断（大規模：101名以上）'
};

const ROUTES = {
  register: register_,
  login: login_,
  changePassword: changePassword_,
  forgotPassword: forgotPassword_,
  logout: logout_,
  me: me_,
  quickDiagnose: quickDiagnose_,
  fullDiagnose: fullDiagnose_,
  history: history_,
  getDiagnosis: getDiagnosis_,
  emailReport: emailReport_
};

/* =========================================================
 * エントリーポイント
 * ======================================================= */
function doGet() {
  return json_({ ok: true, app: APP_NAME, time: nowIso_() });
}

function doPost(e) {
  let res;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = ROUTES[req.action];
    if (!fn) throw err_('BAD', '不明な操作です');
    res = Object.assign({ ok: true }, fn(req) || {});
  } catch (ex) {
    if (!ex.code) console.error(ex && ex.stack ? ex.stack : ex);
    res = { ok: false, code: ex.code || 'ERROR', error: ex.code ? ex.message : 'サーバーエラーが発生しました: ' + ex.message };
  }
  return json_(res);
}

/* =========================================================
 * 初期セットアップ（エディタから1回だけ実行）
 * ======================================================= */
function setup() {
  const ss = getSS_();
  Object.keys(SHEETS).forEach(function (name) {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const h = SHEETS[name];
    sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@'); // 全列テキスト扱い
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#E8EEF7');
    sh.setFrozenRows(1);
  });
  const def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('PEPPER')) p.setProperty('PEPPER', Utilities.getUuid() + Utilities.getUuid());
  getFolder_();
  Logger.log('セットアップ完了。GEMINI_API_KEY が設定済みか確認してください。');
}

/** 期限切れセッションの掃除（任意：時間主導型トリガーで1日1回） */
function cleanupSessions() {
  withLock_(function () {
    const sh = sheet_('sessions');
    const values = sh.getDataRange().getValues();
    const now = new Date();
    for (let i = values.length - 1; i >= 1; i--) {
      if (new Date(values[i][3]) < now) sh.deleteRow(i + 1);
    }
  });
}

/** Gemini接続テスト（任意） */
function testGemini() {
  Logger.log(JSON.stringify(callGemini_('JSONで返答してください。', '{"ok":true} をそのまま返してください。')));
}

/* =========================================================
 * 認証
 * ======================================================= */
function register_(req) {
  const email = normEmail_(req.email);
  const name = str_(req.name, 50);
  const attribute = req.attribute === '法人' ? '法人' : '個人';
  const companyName = attribute === '法人' ? str_(req.companyName, 100) : '';
  const companySize = attribute === '法人' ? String(req.companySize || '') : '';
  if (!name) throw err_('BAD', '氏名を入力してください');
  if (attribute === '法人') {
    if (!companyName) throw err_('BAD', '会社名を入力してください');
    if (SIZES.indexOf(companySize) < 0) throw err_('BAD', '会社規模を選択してください');
  }

  return withLock_(function () {
    if (findUserByEmail_(email)) {
      throw err_('BAD', 'このメールアドレスは登録済みです。パスワードがわからない場合は「パスワードを忘れた方」から再発行してください');
    }
    const temp = tempPassword_();
    const salt = Utilities.getUuid();
    sendTempPasswordMail_(email, name, temp, false); // 送信成功後に保存
    append_('users', {
      id: 'U' + Utilities.getUuid().replace(/-/g, '').slice(0, 16),
      email: email,
      password_hash: hash_(temp, salt),
      salt: salt,
      name: name,
      company_name: companyName,
      attribute: attribute,
      company_size: companySize,
      is_initial_login: 'TRUE',
      created_at: nowIso_(),
      last_login_at: '',
      failed_count: '0',
      locked_until: ''
    });
    return { message: '仮パスワードをメールで送信しました。メールを確認してログインしてください。' };
  });
}

function login_(req) {
  const email = normEmail_(req.email);
  const pw = String(req.password || '');
  return withLock_(function () {
    const u = findUserByEmail_(email);
    if (!u) throw err_('BAD', 'メールアドレスまたはパスワードが正しくありません');
    if (u.locked_until && new Date(u.locked_until) > new Date()) {
      throw err_('LOCKED', 'ログインの失敗が続いたため一時的にロックしています。' + fmtTime_(u.locked_until) +
        ' 以降に再度お試しいただくか、パスワードを再発行してください');
    }
    if (hash_(pw, u.salt) !== u.password_hash) {
      const f = (Number(u.failed_count) || 0) + 1;
      if (f >= MAX_FAILED) {
        update_('users', u._row, {
          failed_count: '0',
          locked_until: new Date(Date.now() + LOCK_MINUTES * 60000).toISOString()
        });
        throw err_('LOCKED', MAX_FAILED + '回続けて失敗したため、' + LOCK_MINUTES + '分間ログインできません');
      }
      update_('users', u._row, { failed_count: String(f) });
      throw err_('BAD', 'メールアドレスまたはパスワードが正しくありません（あと' + (MAX_FAILED - f) + '回失敗するとロックされます）');
    }
    update_('users', u._row, { failed_count: '0', locked_until: '', last_login_at: nowIso_() });
    return { token: createSession_(u.id), user: publicUser_(u) };
  });
}

function changePassword_(req) {
  return withLock_(function () {
    const u = auth_(req, true);
    const cur = String(req.currentPassword || '');
    const nw = String(req.newPassword || '');
    if (hash_(cur, u.salt) !== u.password_hash) throw err_('BAD', '現在のパスワード（仮パスワード）が正しくありません');
    validatePassword_(nw);
    if (nw === cur) throw err_('BAD', '現在と異なるパスワードを設定してください');
    const salt = Utilities.getUuid();
    update_('users', u._row, { password_hash: hash_(nw, salt), salt: salt, is_initial_login: 'FALSE' });
    u.is_initial_login = 'FALSE';
    return { message: 'パスワードを設定しました', user: publicUser_(u) };
  });
}

function forgotPassword_(req) {
  const email = normEmail_(req.email);
  const msg = { message: '登録済みのメールアドレスであれば、新しい仮パスワードを送信しました。届かない場合は5分ほど待ってから再度お試しください。' };
  const cache = CacheService.getScriptCache();
  if (cache.get('fp:' + email)) return msg; // 連続送信の抑止
  cache.put('fp:' + email, '1', 300);

  return withLock_(function () {
    const u = findUserByEmail_(email);
    if (!u) return msg;
    const temp = tempPassword_();
    const salt = Utilities.getUuid();
    sendTempPasswordMail_(u.email, u.name, temp, true);
    update_('users', u._row, {
      password_hash: hash_(temp, salt), salt: salt, is_initial_login: 'TRUE', failed_count: '0', locked_until: ''
    });
    removeSessionsOfUser_(u.id);
    return msg;
  });
}

function logout_(req) {
  const token = String(req.token || '');
  if (token) withLock_(function () { removeSession_(token); });
  return { message: 'ログアウトしました' };
}

function me_(req) {
  return { user: publicUser_(auth_(req, true)) };
}

/* =========================================================
 * 診断
 * ======================================================= */
function quickDiagnose_(req) {
  const answers = cleanAnswers_(req.answers);
  let userId = 'guest';
  if (req.token) {
    try { userId = auth_(req, true).id; } catch (e) { userId = 'guest'; }
  }
  if (userId === 'guest') checkGuestQuota_();
  const result = normalizeQuick_(callGemini_(SYSTEM_PROMPT, buildQuickPrompt_(answers)));
  const id = saveDiagnosis_(userId, '簡易', answers, result);
  return { diagnosisId: id, type: '簡易', createdAt: nowIso_(), result: result };
}

function fullDiagnose_(req) {
  const u = auth_(req, false);
  const type = String(req.type || '');
  if (FULL_TYPES.indexOf(type) < 0) throw err_('BAD', '診断タイプが正しくありません');
  const answers = cleanAnswers_(req.answers);
  const result = normalizeFull_(callGemini_(SYSTEM_PROMPT, buildFullPrompt_(type, answers, u)));
  const id = saveDiagnosis_(u.id, type, answers, result);
  return { diagnosisId: id, type: type, createdAt: nowIso_(), result: result };
}

function history_(req) {
  const u = auth_(req, false);
  const list = rows_('diagnostics')
    .filter(function (d) { return d.user_id === u.id; })
    .map(function (d) {
      let r = {};
      try { r = JSON.parse(d.result_json); } catch (e) {}
      return { diagnosisId: d.diagnosis_id, type: d.type, createdAt: d.created_at, level: r.level || '', score: r.score || 0 };
    })
    .sort(function (a, b) { return new Date(b.createdAt) - new Date(a.createdAt); })
    .slice(0, 50);
  return { history: list };
}

function getDiagnosis_(req) {
  const u = auth_(req, false);
  const d = findDiagnosis_(req.diagnosisId, u.id);
  return {
    diagnosisId: d.diagnosis_id, type: d.type, createdAt: d.created_at,
    answers: JSON.parse(d.answers_json || '[]'), result: JSON.parse(d.result_json || '{}')
  };
}

function emailReport_(req) {
  const u = auth_(req, false);
  const d = findDiagnosis_(req.diagnosisId, u.id);
  const result = JSON.parse(d.result_json || '{}');
  const stamp = Utilities.formatDate(new Date(d.created_at), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
  const pdf = Utilities.newBlob(reportHtml_(d.type, d.created_at, result, u), 'text/html', 'report.html')
    .getAs('application/pdf')
    .setName('AI活用診断レポート_' + stamp + '.pdf');

  const file = getFolder_().createFile(pdf);
  withLock_(function () { update_('diagnostics', d._row, { pdf_file_id: file.getId() }); });

  MailApp.sendEmail({
    to: u.email,
    subject: '[' + APP_NAME + '] 診断レポート（' + (TYPE_LABEL[d.type] || d.type) + '）',
    body: u.name + ' 様\n\n' + APP_NAME + 'をご利用いただきありがとうございます。\n診断レポートをPDFでお送りします。\n\n' +
      (prop_('APP_URL') ? 'マイページ: ' + prop_('APP_URL') + '\n\n' : '') + '※ このメールは送信専用です。',
    name: APP_NAME,
    attachments: [pdf]
  });
  return { message: 'レポートのPDFを ' + u.email + ' に送信しました' };
}

/* =========================================================
 * Gemini プロンプト
 * ======================================================= */
const SYSTEM_PROMPT = [
  'あなたは個人と企業の生成AI活用を支援する、経験豊富で中立的なコンサルタントです。',
  '# ルール',
  '- 出力は指定されたJSONオブジェクトのみ。前後の文章やコードブロック記号は付けない。',
  '- 日本語で、専門用語には短い補足を添える。読み手はAIに詳しくない人も含む。',
  '- 推奨ツールは実在し広く使われているものに限る。料金や効果は断定せず「目安」として幅で示す。',
  '- 回答から読み取れない前提を置いた場合は、その前提を明記する。',
  '- 誇張や不安をあおる表現は避け、現実的で今日から実行できる提案を優先する。',
  '- 個人情報・機密情報・著作権・AIの誤回答（ハルシネーション）への配慮を常に含める。'
].join('\n');

const LEVEL_RULE = [
  '# レベル基準',
  '- 初級: 生成AIをほとんど使っていない、または使い方が定まっていない',
  '- 中級: 日常的に使い、用途がある程度定まっている',
  '- 上級: 指示の工夫や結果の検証ができ、業務や生活の仕組みに組み込んでいる',
  'score は 0〜100 の整数。初級 0〜39、中級 40〜74、上級 75〜100 の範囲で level と一致させる。'
].join('\n');

const TYPE_FOCUS = {
  '個人': '個人の生産性・学習・収入づくりの観点で診断する。目的（業務効率化／自己学習／副業・フリーランス）に合わせ、ROIは「月あたりの節約時間」「月額費用」「元が取れるまでの目安」で示す。',
  '法人_小規模': '従業員1〜20名の会社。IT専任者がいない前提で、低コストですぐ使える汎用ツール中心に、経営者主導で小さく始める手順を示す。ROIは人件費換算の削減効果で示す。',
  '法人_中規模': '従業員21〜100名の会社。部門単位での試行から全社展開、利用ガイドライン整備、社員教育、効果測定の仕組みを含める。ROIは部門ごとの工数削減と費用で示す。',
  '法人_大規模': '従業員101名以上の会社。全社ガバナンス（承認フロー・監査）、情報セキュリティ・法務、既存システム連携、推進体制、段階的な全社展開を含める。ROIは全社と先行部門の両面で示す。'
};

function answersText_(answers) {
  return answers.map(function (x, i) { return 'Q' + (i + 1) + '. ' + x.q + '\nA. ' + x.a; }).join('\n\n');
}

function buildQuickPrompt_(answers) {
  return [
    '# 依頼',
    '簡易診断です。以下の回答から回答者のAI活用レベルを判定し、最初の一歩となる具体的な活用法を1つ提案してください。',
    '',
    LEVEL_RULE,
    '',
    '# 回答',
    answersText_(answers),
    '',
    '# 出力JSONの形式',
    JSON.stringify({
      level: '初級|中級|上級',
      score: 0,
      summary: '現状の評価（2〜3文）',
      firstStep: {
        title: 'おすすめの最初の一歩（20字程度）',
        description: '具体的なやり方（3〜4文。例となる指示文を1つ含める）',
        tools: ['ツール名']
      },
      tips: ['短い助言（3つ）']
    }, null, 2)
  ].join('\n');
}

function buildFullPrompt_(type, answers, u) {
  const profile = u.attribute === '法人'
    ? '属性: 法人／登録時の会社規模: ' + (u.company_size || '不明')
    : '属性: 個人';
  return [
    '# 依頼',
    '本診断です。以下の回答をもとに、詳細な診断レポートを作成してください。',
    '',
    '# 診断タイプ',
    TYPE_LABEL[type],
    TYPE_FOCUS[type],
    profile,
    '',
    LEVEL_RULE,
    '',
    '# 回答',
    answersText_(answers),
    '',
    '# 出力JSONの形式（各配列の個数を守る）',
    JSON.stringify({
      level: '初級|中級|上級',
      score: 0,
      summary: '総評（4〜6文）',
      strengths: ['強み（2〜3個）'],
      challenges: ['課題（2〜3個）'],
      roi: {
        headline: 'ROI見込みの要点（1〜2文）',
        timeSaved: '削減できる時間の目安（例: 月20〜40時間）',
        monthlyCost: '月額費用の目安',
        payback: '費用回収までの目安',
        assumptions: ['試算の前提（2〜3個）']
      },
      recommendedTools: [
        { name: 'ツール名', category: '分類', purpose: 'この回答者での使いどころ', price: '料金の目安' }
      ],
      roadmap: [
        { phase: '段階名', period: '期間の目安', goal: 'この段階のゴール', actions: ['具体的な行動（2〜4個）'] }
      ],
      security: [
        { risk: '想定されるリスク', measure: '具体的な対策' }
      ],
      nextActions: ['今週中にできる行動（3個）']
    }, null, 2),
    '',
    'recommendedTools は3〜6件、roadmap は3〜4段階、security は3〜5件。'
  ].join('\n');
}

/* =========================================================
 * Gemini API 呼び出し
 * ======================================================= */
function callGemini_(system, prompt) {
  const key = prop_('GEMINI_API_KEY');
  if (!key) throw err_('CONFIG', 'サーバー設定エラー: GEMINI_API_KEY が未設定です');
  const model = prop_('GEMINI_MODEL', 'gemini-2.5-flash');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
  const payload = {
    systemInstruction: { parts: [{ text: system }] },
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.6, responseMimeType: 'application/json', maxOutputTokens: 16384 }
  };

  for (let i = 0; i < 3; i++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code === 200) {
      const data = JSON.parse(res.getContentText());
      const cand = (data.candidates || [])[0] || {};
      const parts = (cand.content && cand.content.parts) || [];
      const text = parts.filter(function (p) { return !p.thought; }).map(function (p) { return p.text || ''; }).join('');
      if (!text) throw err_('AI', 'AIから回答が得られませんでした。もう一度お試しください');
      return parseJson_(text);
    }
    if ([429, 500, 503].indexOf(code) >= 0 && i < 2) {
      Utilities.sleep(2000 * (i + 1));
      continue;
    }
    console.error('Gemini error', code, res.getContentText());
    throw err_('AI', code === 429
      ? 'AIが混み合っています。少し時間をおいて再度お試しください'
      : 'AI診断の作成に失敗しました（エラー ' + code + '）');
  }
}

function parseJson_(text) {
  let t = String(text).replace(/```json|```/g, '').trim();
  const s = t.indexOf('{'), e = t.lastIndexOf('}');
  if (s >= 0 && e > s) t = t.slice(s, e + 1);
  try {
    return JSON.parse(t);
  } catch (ex) {
    console.error('JSON parse failed', text);
    throw err_('AI', 'AIの回答を読み取れませんでした。もう一度お試しください');
  }
}

/* =========================================================
 * 結果の整形（欠損に強くする）
 * ======================================================= */
function s_(v) { return v == null ? '' : String(v); }
function arr_(v) { return Array.isArray(v) ? v : []; }
function level_(v) { return ['初級', '中級', '上級'].indexOf(v) >= 0 ? v : '初級'; }
function score_(v) { const n = Math.round(Number(v)); return isNaN(n) ? 0 : Math.max(0, Math.min(100, n)); }

function normalizeQuick_(r) {
  r = r || {};
  const f = r.firstStep || {};
  return {
    level: level_(r.level),
    score: score_(r.score),
    summary: s_(r.summary),
    firstStep: { title: s_(f.title), description: s_(f.description), tools: arr_(f.tools).map(s_) },
    tips: arr_(r.tips).map(s_)
  };
}

function normalizeFull_(r) {
  r = r || {};
  const roi = r.roi || {};
  return {
    level: level_(r.level),
    score: score_(r.score),
    summary: s_(r.summary),
    strengths: arr_(r.strengths).map(s_),
    challenges: arr_(r.challenges).map(s_),
    roi: {
      headline: s_(roi.headline), timeSaved: s_(roi.timeSaved), monthlyCost: s_(roi.monthlyCost),
      payback: s_(roi.payback), assumptions: arr_(roi.assumptions).map(s_)
    },
    recommendedTools: arr_(r.recommendedTools).map(function (t) {
      t = t || {};
      return { name: s_(t.name), category: s_(t.category), purpose: s_(t.purpose), price: s_(t.price) };
    }),
    roadmap: arr_(r.roadmap).map(function (p) {
      p = p || {};
      return { phase: s_(p.phase), period: s_(p.period), goal: s_(p.goal), actions: arr_(p.actions).map(s_) };
    }),
    security: arr_(r.security).map(function (x) {
      x = x || {};
      return { risk: s_(x.risk), measure: s_(x.measure) };
    }),
    nextActions: arr_(r.nextActions).map(s_)
  };
}

/* =========================================================
 * PDF用HTML
 * ======================================================= */
function reportHtml_(type, createdAt, r, u) {
  const e = escHtml_;
  const li = function (a) { return '<ul>' + arr_(a).map(function (x) { return '<li>' + e(x) + '</li>'; }).join('') + '</ul>'; };
  const date = Utilities.formatDate(new Date(createdAt), 'Asia/Tokyo', 'yyyy年M月d日 HH:mm');
  let h = '<html><head><meta charset="UTF-8"><style>' +
    'body{font-family:sans-serif;color:#14213D;font-size:11pt;line-height:1.7;}' +
    'h1{font-size:18pt;margin:0 0 4px;}h2{font-size:13pt;border-bottom:2px solid #2F5BEA;padding-bottom:3px;margin-top:22px;}' +
    '.meta{color:#5B6B82;font-size:9pt;}.lv{font-size:22pt;font-weight:bold;color:#2F5BEA;}' +
    'table{border-collapse:collapse;width:100%;}td,th{border:1px solid #DCE3EC;padding:6px;vertical-align:top;font-size:10pt;}' +
    'th{background:#EEF2F8;text-align:left;}.box{background:#F3F6F9;padding:10px;border-radius:6px;}' +
    '</style></head><body>';
  h += '<h1>' + APP_NAME + ' レポート</h1>';
  h += '<div class="meta">' + e(TYPE_LABEL[type] || type) + '　' + date + '　' + e(u.attribute === '法人' ? u.company_name : u.name) + '</div>';
  h += '<h2>判定結果</h2><div class="lv">' + e(r.level) + '（' + e(r.score) + '点）</div><p>' + e(r.summary) + '</p>';

  if (r.firstStep) {
    h += '<h2>最初の一歩：' + e(r.firstStep.title) + '</h2><p>' + e(r.firstStep.description) + '</p>';
    if (r.firstStep.tools && r.firstStep.tools.length) h += '<p>おすすめツール：' + e(r.firstStep.tools.join('、')) + '</p>';
    if (r.tips && r.tips.length) h += '<h2>ヒント</h2>' + li(r.tips);
  }
  if (r.strengths) h += '<h2>強み</h2>' + li(r.strengths) + '<h2>課題</h2>' + li(r.challenges);
  if (r.roi) {
    h += '<h2>ROIの見込み</h2><div class="box"><p>' + e(r.roi.headline) + '</p>' +
      '<table><tr><th>削減できる時間</th><th>月額費用</th><th>費用回収</th></tr><tr><td>' + e(r.roi.timeSaved) +
      '</td><td>' + e(r.roi.monthlyCost) + '</td><td>' + e(r.roi.payback) + '</td></tr></table>' +
      '<p class="meta">試算の前提</p>' + li(r.roi.assumptions) + '</div>';
  }
  if (r.recommendedTools && r.recommendedTools.length) {
    h += '<h2>推奨AIツール</h2><table><tr><th>ツール</th><th>分類</th><th>使いどころ</th><th>料金の目安</th></tr>' +
      r.recommendedTools.map(function (t) {
        return '<tr><td>' + e(t.name) + '</td><td>' + e(t.category) + '</td><td>' + e(t.purpose) + '</td><td>' + e(t.price) + '</td></tr>';
      }).join('') + '</table>';
  }
  if (r.roadmap && r.roadmap.length) {
    h += '<h2>導入ステップ</h2>' + r.roadmap.map(function (p, i) {
      return '<p><b>ステップ' + (i + 1) + '　' + e(p.phase) + '</b>（' + e(p.period) + '）<br>ゴール：' + e(p.goal) + '</p>' + li(p.actions);
    }).join('');
  }
  if (r.security && r.security.length) {
    h += '<h2>セキュリティの留意点</h2><table><tr><th>リスク</th><th>対策</th></tr>' +
      r.security.map(function (x) { return '<tr><td>' + e(x.risk) + '</td><td>' + e(x.measure) + '</td></tr>'; }).join('') + '</table>';
  }
  if (r.nextActions && r.nextActions.length) h += '<h2>今週やること</h2>' + li(r.nextActions);
  h += '<p class="meta" style="margin-top:24px;">※ 本レポートはAIによる診断です。料金や効果は目安であり、導入時は最新情報をご確認ください。</p>';
  return h + '</body></html>';
}

/* =========================================================
 * セッション
 * ======================================================= */
function createSession_(userId) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const exp = new Date(Date.now() + SESSION_DAYS * 86400000).toISOString();
  append_('sessions', { token: token, user_id: userId, created_at: nowIso_(), expires_at: exp });
  CacheService.getScriptCache().put('s:' + token, JSON.stringify({ u: userId, e: exp }), 21600);
  return token;
}

function auth_(req, allowInitial) {
  const token = String(req.token || '');
  if (!token) throw err_('AUTH', 'ログインしてください');
  const cache = CacheService.getScriptCache();
  let userId, exp;
  const c = cache.get('s:' + token);
  if (c) {
    const o = JSON.parse(c);
    userId = o.u; exp = o.e;
  } else {
    const row = rows_('sessions').find(function (r) { return r.token === token; });
    if (!row) throw err_('AUTH', 'ログインの有効期限が切れました。もう一度ログインしてください');
    userId = row.user_id; exp = row.expires_at;
    cache.put('s:' + token, JSON.stringify({ u: userId, e: exp }), 21600);
  }
  if (new Date(exp) < new Date()) {
    withLock_(function () { removeSession_(token); });
    throw err_('AUTH', 'ログインの有効期限が切れました。もう一度ログインしてください');
  }
  const u = rows_('users').find(function (x) { return x.id === userId; });
  if (!u) throw err_('AUTH', 'ログインしてください');
  if (!allowInitial && u.is_initial_login === 'TRUE') throw err_('MUST_CHANGE', '先にパスワードの本設定を完了してください');
  return u;
}

function removeSession_(token) {
  CacheService.getScriptCache().remove('s:' + token);
  const sh = sheet_('sessions');
  const v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (v[i][0] === token) sh.deleteRow(i + 1);
  }
}

function removeSessionsOfUser_(userId) {
  const cache = CacheService.getScriptCache();
  const sh = sheet_('sessions');
  const v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (v[i][1] === userId) {
      cache.remove('s:' + v[i][0]);
      sh.deleteRow(i + 1);
    }
  }
}

/* =========================================================
 * ユーティリティ
 * ======================================================= */
let LOCK_HELD_ = false;
function withLock_(fn) {
  if (LOCK_HELD_) return fn();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  LOCK_HELD_ = true;
  try { return fn(); } finally { LOCK_HELD_ = false; lock.releaseLock(); }
}

function err_(code, message) { const e = new Error(message); e.code = code; return e; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function prop_(k, def) { const v = PropertiesService.getScriptProperties().getProperty(k); return v == null || v === '' ? (def || '') : v; }
function nowIso_() { return new Date().toISOString(); }
function fmtTime_(iso) { return Utilities.formatDate(new Date(iso), 'Asia/Tokyo', 'HH:mm'); }
function str_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }

function getSS_() {
  const id = prop_('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw err_('CONFIG', 'サーバー設定エラー: SPREADSHEET_ID が未設定です');
  return ss;
}
function sheet_(name) { return getSS_().getSheetByName(name); }

function rows_(name) {
  const v = sheet_(name).getDataRange().getValues();
  const h = v.shift() || [];
  return v.map(function (r, i) {
    const o = { _row: i + 2 };
    h.forEach(function (k, j) { o[k] = r[j] === '' || r[j] == null ? '' : String(r[j]); });
    return o;
  });
}
function append_(name, obj) {
  sheet_(name).appendRow(SHEETS[name].map(function (k) { return obj[k] == null ? '' : String(obj[k]); }));
}
function update_(name, row, obj) {
  const sh = sheet_(name);
  const h = SHEETS[name];
  Object.keys(obj).forEach(function (k) {
    const c = h.indexOf(k);
    if (c >= 0) sh.getRange(row, c + 1).setValue(String(obj[k]));
  });
}

function normEmail_(v) {
  const e = String(v || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e) || e.length > 254) throw err_('BAD', 'メールアドレスの形式が正しくありません');
  return e;
}
function findUserByEmail_(email) { return rows_('users').find(function (u) { return u.email === email; }); }
function publicUser_(u) {
  return {
    id: u.id, email: u.email, name: u.name, companyName: u.company_name,
    attribute: u.attribute, companySize: u.company_size, mustChangePassword: u.is_initial_login === 'TRUE'
  };
}

function toHex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join(''); }
function hash_(pw, salt) {
  const pepper = prop_('PEPPER');
  let h = salt + ':' + pw + ':' + pepper;
  for (let i = 0; i < HASH_ITER; i++) {
    h = toHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  }
  return h;
}
function tempPassword_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now() + Math.random());
  let p = '';
  for (let i = 0; i < 10; i++) p += chars[(bytes[i] & 0xff) % chars.length];
  return p.replace(/^(.{9})./, '$1' + '23456789'[(bytes[20] & 0xff) % 8]); // 数字を最低1文字含める
}
function validatePassword_(pw) {
  if (pw.length < 8 || pw.length > 64) throw err_('BAD', 'パスワードは8〜64文字で設定してください');
  if (!/[A-Za-z]/.test(pw) || !/\d/.test(pw)) throw err_('BAD', 'パスワードには英字と数字の両方を含めてください');
}

function sendTempPasswordMail_(to, name, temp, isReset) {
  const url = prop_('APP_URL');
  MailApp.sendEmail({
    to: to,
    subject: '[' + APP_NAME + '] ' + (isReset ? '仮パスワード再発行' : '会員登録') + 'のお知らせ',
    body: name + ' 様\n\n' +
      (isReset ? '仮パスワードを再発行しました。' : APP_NAME + 'へのご登録ありがとうございます。') + '\n\n' +
      '仮パスワード: ' + temp + '\n\n' +
      'ログイン後、新しいパスワードの設定画面が表示されます。\n' +
      (url ? 'ログイン: ' + url + '\n' : '') +
      '\n心当たりがない場合は、このメールを破棄してください。\n※ このメールは送信専用です。',
    name: APP_NAME
  });
}

function cleanAnswers_(list) {
  const a = arr_(list).slice(0, 12).map(function (x) {
    x = x || {};
    return { q: str_(x.q, 200), a: str_(x.a, 500) || '（未回答）' };
  }).filter(function (x) { return x.q; });
  if (a.length < 3) throw err_('BAD', '回答が不足しています');
  return a;
}

function checkGuestQuota_() {
  const limit = Number(prop_('GUEST_DAILY_LIMIT', '200'));
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  withLock_(function () {
    const p = PropertiesService.getScriptProperties();
    let q = {};
    try { q = JSON.parse(p.getProperty('GUEST_QUOTA') || '{}'); } catch (e) {}
    const n = q.date === today ? Number(q.count) || 0 : 0;
    if (n >= limit) throw err_('LIMIT', '本日の簡易診断の受付数が上限に達しました。明日以降にお試しいただくか、会員登録して本診断をご利用ください');
    p.setProperty('GUEST_QUOTA', JSON.stringify({ date: today, count: n + 1 }));
  });
}

function saveDiagnosis_(userId, type, answers, result) {
  const id = 'D' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + Utilities.getUuid().slice(0, 6);
  withLock_(function () {
    append_('diagnostics', {
      diagnosis_id: id, user_id: userId, type: type,
      answers_json: JSON.stringify(answers), result_json: JSON.stringify(result),
      created_at: nowIso_(), pdf_file_id: ''
    });
  });
  return id;
}

function findDiagnosis_(id, userId) {
  const d = rows_('diagnostics').find(function (x) { return x.diagnosis_id === String(id || ''); });
  if (!d || d.user_id !== userId) throw err_('BAD', '診断結果が見つかりません');
  return d;
}

function getFolder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('DRIVE_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  const f = DriveApp.createFolder('AI活用診断レポート');
  p.setProperty('DRIVE_FOLDER_ID', f.getId());
  return f;
}

function escHtml_(v) {
  return String(v == null ? '' : v).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
