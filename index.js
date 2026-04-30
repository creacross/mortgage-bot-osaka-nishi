// ============================================================
// 大阪西区特化 住宅ローン相談 LINE Bot
// Node.js / Render.com版
// ============================================================

const express = require('express');
const { google } = require('googleapis');
const app = express();

// ── 環境変数 ──
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN || '';
const LINE_CHANNEL_SECRET       = process.env.LINE_CHANNEL_SECRET || '';
const SPREADSHEET_ID            = process.env.SPREADSHEET_ID || '';
const PORT                      = process.env.PORT || 3000;

// ── Google Sheets 認証 ──
let sheets = null;
try {
  const sa = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON || '{}');
  const auth = new google.auth.GoogleAuth({
    credentials: sa,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  sheets = google.sheets({ version: 'v4', auth });
  console.log('Google Sheets: 認証成功');
} catch (err) {
  console.error('Google Sheets: 認証失敗', err.message);
}

app.use(express.json({ verify: (req, res, buf) => { req.rawBody = buf; } }));

// ── ユーザーステート管理 ──
const userStates = {};

function getState(userId) {
  return userStates[userId] || { step: 'NONE', answers: {}, mode: 'bot' };
}
function setState(userId, state) { userStates[userId] = state; }
function clearState(userId) {
  const s = userStates[userId];
  userStates[userId] = { step: 'NONE', answers: {}, mode: s ? s.mode : 'bot' };
}
function setManual(userId) {
  const s = getState(userId);
  s.mode = 'manual'; s.manualGreeted = false;
  s.step = 'NONE'; s.answers = {};
  setState(userId, s);
}
function setBot(userId) {
  const s = getState(userId);
  s.mode = 'bot'; s.step = 'NONE'; s.answers = {};
  setState(userId, s);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 返済シミュレーション計算
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function calcMonthly(principal, annualRate, years) {
  if (annualRate === 0) return Math.round(principal / (years * 12));
  const r = annualRate / 100 / 12;
  const n = years * 12;
  return Math.round(principal * r * Math.pow(1 + r, n) / (Math.pow(1 + r, n) - 1));
}

function formatYen(amount) {
  return amount.toLocaleString('ja-JP') + '円';
}

function generateSimulation(answers) {
  // 物件価格の中央値を使用
  const priceMap = {
    '〜3000万円': 2500, '3000〜5000万円': 4000,
    '5000〜8000万円': 6500, '8000万円以上': 10000,
    'まだわからない': 4000
  };
  const selfFundMap = {
    'なし（フルローン希望）': 0, '〜200万円': 100,
    '200〜500万円': 350, '500万円以上': 700
  };
  const incomeMap = {
    '〜400万円': 350, '400〜600万円': 500,
    '600〜800万円': 700, '800〜1000万円': 900, '1000万円以上': 1200
  };

  const price    = (priceMap[answers.price] || 4000) * 10000;
  const selfFund = (selfFundMap[answers.selfFund] || 0) * 10000;
  const loan     = price - selfFund;
  const income   = (incomeMap[answers.income] || 500) * 10000;
  const years    = 35;

  const v1 = calcMonthly(loan, 0.5,  years); // 変動0.5%
  const v2 = calcMonthly(loan, 1.5,  years); // 固定10年1.5%
  const v3 = calcMonthly(loan, 2.0,  years); // 全期間固定2.0%

  const ratio = Math.round((v1 * 12) / income * 100);

  let ratioComment = '';
  let ratioEmoji   = '';
  if (ratio <= 25) {
    ratioComment = '返済余裕度が高く、審査通過の可能性が高い水準です。';
    ratioEmoji   = '✅';
  } else if (ratio <= 35) {
    ratioComment = '一般的な返済比率の範囲内です。';
    ratioEmoji   = '🟡';
  } else {
    ratioComment = '返済比率がやや高めです。頭金の増額や借入額の見直しをご検討ください。';
    ratioEmoji   = '⚠️';
  }

  // 雇用形態・属性に応じた銀行タイプの案内
  let bankAdvice = '';
  if (answers.employment === '会社員' || answers.employment === '公務員') {
    bankAdvice = '👉 ネット銀行（変動金利が低い）または大手行が向いています。';
  } else if (answers.employment === '自営業・フリーランス') {
    bankAdvice = '👉 信用金庫や地方銀行が柔軟に対応しやすい傾向があります。確定申告3年分の準備が重要です。';
  } else {
    bankAdvice = '👉 まずは在籍期間と収入の安定性を重視した金融機関への相談をおすすめします。';
  }

  return {
    loanAmount: loan,
    monthly_v1: v1,
    monthly_v2: v2,
    monthly_v3: v3,
    ratio,
    ratioComment,
    ratioEmoji,
    bankAdvice,
    message: [
      `📊 返済シミュレーション結果`,
      ``,
      `▼ 借入想定額：${formatYen(loan)}`,
      ``,
      `【変動金利 0.5%】`,
      `　毎月：${formatYen(v1)}`,
      ``,
      `【固定10年 1.5%】`,
      `　毎月：${formatYen(v2)}`,
      ``,
      `【全期間固定 2.0%】`,
      `　毎月：${formatYen(v3)}`,
      ``,
      `▼ 返済比率（変動基準）：${ratio}%`,
      `${ratioEmoji} ${ratioComment}`,
      ``,
      `${bankAdvice}`,
      ``,
      `※ あくまで参考値です。実際の審査結果を保証するものではありません。`
    ].join('\n')
  };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// スプレッドシート書き込み
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async function writeToSheet(data) {
  if (!sheets || !SPREADSHEET_ID) return;
  try {
    const ss = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const names = ss.data.sheets.map(s => s.properties.title);
    const SHEET = 'ローン相談';

    if (!names.includes(SHEET)) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { requests: [{ addSheet: { properties: { title: SHEET } } }] }
      });
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET}!A1`,
        valueInputOption: 'RAW',
        requestBody: { values: [[
          '受付日時','LINE ID','お名前','相談種別','エリア','物件種別',
          '物件価格','年収帯','雇用形態','自己資金','他ローン',
          '検討時期','借入想定額','月返済(変動)','返済比率','ステータス'
        ]]}
      });
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: [[
        data.timestamp, data.userId, data.name, data.consultType,
        data.area, data.propertyType, data.price, data.income,
        data.employment, data.selfFund, data.otherLoan,
        data.timing, data.loanAmount, data.monthlyPayment,
        data.ratio + '%', '未対応'
      ]]}
    });
    console.log('スプレッドシート記録完了');
  } catch (err) {
    console.error('スプレッドシートエラー:', err.message);
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ヘルスチェック
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
app.get('/', (req, res) => res.status(200).send('Mortgage Bot is running.'));

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Webhook
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
app.post('/webhook', async (req, res) => {
  res.status(200).json({ status: 'ok' });
  const events = req.body.events || [];

  for (const event of events) {
    try {
      const userId = event.source?.userId;
      if (!userId) continue;

      if (event.type === 'follow') { await handleFollow(event); continue; }

      if (event.type === 'message' && event.message.type === 'text') {
        const text = event.message.text.trim();

        // 担当者コマンド
        if (text === '#対応開始') {
          setManual(userId);
          await reply(event.replyToken, [{ type: 'text', text: '担当スタッフが対応いたします。どうぞお気軽にご質問ください。' }]);
          continue;
        }
        if (text === '#bot再開') {
          setBot(userId);
          await reply(event.replyToken, [{ type: 'text', text: 'AIアシスタントが対応を再開しました。\n「相談したい」と送っていただければご案内を開始します。' }]);
          continue;
        }
        if (text === '#状態確認') {
          const s = getState(userId);
          await reply(event.replyToken, [{ type: 'text', text: `現在のモード：${s.mode === 'manual' ? '手動対応中' : 'Bot対応中'}` }]);
          continue;
        }

        // 「相談したい」でリセット
        if (text === '相談したい' || text === '相談') {
          setState(userId, { step: 'SELECT_TYPE', answers: {}, mode: 'bot' });
          await sendTypeSelection(event.replyToken);
          continue;
        }

        // 手動モード
        const s = getState(userId);
        if (s.mode === 'manual') {
          if (!s.manualGreeted) {
            s.manualGreeted = true;
            setState(userId, s);
            await reply(event.replyToken, [{ type: 'text', text: '担当スタッフが対応いたします。\n他にも聞きたいことがあればお気軽にご質問ください。' }]);
          }
          continue;
        }

        await handleMessage(event);
        continue;
      }

      if (event.type === 'postback') {
        const s = getState(userId);
        if (s.mode === 'manual') continue;
        await handlePostback(event);
      }
    } catch (err) {
      console.error('Event error:', err);
    }
  }
});

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 友だち追加
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async function handleFollow(event) {
  const userId = event.source.userId;
  setState(userId, { step: 'SELECT_TYPE', answers: {}, mode: 'bot' });

  await reply(event.replyToken, [
    {
      type: 'text',
      text: '友だち追加ありがとうございます！\n\n大阪市西区を中心に地域密着で\n住宅ローンのご相談を承っております。\n\n📍 クレアクロス株式会社\n\n気軽にご相談ください。\n個人情報は厳重に管理し、\nご同意なく第三者に提供することはありません。'
    },
    {
      type: 'template',
      altText: 'ご相談の種類を選んでください',
      template: {
        type: 'buttons',
        title: 'どのようなご相談ですか？',
        text: '当てはまるものをお選びください',
        actions: [
          { type: 'postback', label: '🏠 マイホームを買いたい', data: 'type=マイホーム' },
          { type: 'postback', label: '📊 返済額を試算したい', data: 'type=試算のみ' },
          { type: 'postback', label: '🔄 ローンを借り換えたい', data: 'type=借り換え' },
          { type: 'postback', label: '💼 投資物件を検討中', data: 'type=投資' }
        ]
      }
    }
  ]);
}

async function sendTypeSelection(replyToken) {
  await reply(replyToken, [
    {
      type: 'template',
      altText: 'ご相談の種類を選んでください',
      template: {
        type: 'buttons',
        title: 'どのようなご相談ですか？',
        text: '当てはまるものをお選びください',
        actions: [
          { type: 'postback', label: '🏠 マイホームを買いたい', data: 'type=マイホーム' },
          { type: 'postback', label: '📊 返済額を試算したい', data: 'type=試算のみ' },
          { type: 'postback', label: '🔄 ローンを借り換えたい', data: 'type=借り換え' },
          { type: 'postback', label: '💼 投資物件を検討中', data: 'type=投資' }
        ]
      }
    }
  ]);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Postback処理
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async function handlePostback(event) {
  const userId = event.source.userId;
  const params = parseData(event.postback.data);
  const state  = getState(userId);

  // 相談種別の選択
  if (params.type) {
    state.answers.consultType = params.type;
    state.step = 'ASK_AREA';
    setState(userId, state);
    await reply(event.replyToken, [{
      type: 'text',
      text: `${params.type}のご相談ですね。ありがとうございます！\n\nご検討のエリアを教えてください。\n（例：大阪市西区、南堀江、堀江、靱公園 など）`
    }]);
    return;
  }

  // 物件種別
  if (params.propertyType) {
    state.answers.propertyType = params.propertyType;
    state.step = 'ASK_PRICE';
    setState(userId, state);
    await reply(event.replyToken, [{
      type: 'template',
      altText: '物件価格帯を選んでください',
      template: {
        type: 'buttons',
        title: 'ご検討の物件価格は？',
        text: '目安でOKです',
        actions: [
          { type: 'postback', label: '〜3,000万円',      data: 'price=〜3000万円' },
          { type: 'postback', label: '3,000〜5,000万円', data: 'price=3000〜5000万円' },
          { type: 'postback', label: '5,000〜8,000万円', data: 'price=5000〜8000万円' },
          { type: 'postback', label: '8,000万円以上',    data: 'price=8000万円以上' }
        ]
      }
    }]);
    return;
  }

  // 物件価格
  if (params.price) {
    state.answers.price = params.price;
    state.step = 'ASK_INCOME';
    setState(userId, state);
    await reply(event.replyToken, [{
      type: 'template',
      altText: '世帯年収を選んでください',
      template: {
        type: 'buttons',
        title: '世帯年収の目安は？',
        text: '審査可能性の確認に使用します',
        actions: [
          { type: 'postback', label: '〜400万円',        data: 'income=〜400万円' },
          { type: 'postback', label: '400〜600万円',     data: 'income=400〜600万円' },
          { type: 'postback', label: '600〜1,000万円',   data: 'income=600〜1000万円' },
          { type: 'postback', label: '1,000万円以上',    data: 'income=1000万円以上' }
        ]
      }
    }]);
    return;
  }

  // 年収
  if (params.income) {
    state.answers.income = params.income;
    state.step = 'ASK_EMPLOYMENT';
    setState(userId, state);
    await reply(event.replyToken, [{
      type: 'template',
      altText: 'ご職業を選んでください',
      template: {
        type: 'buttons',
        title: 'ご職業を教えてください',
        text: '審査の参考にします',
        actions: [
          { type: 'postback', label: '会社員',             data: 'employment=会社員' },
          { type: 'postback', label: '公務員',             data: 'employment=公務員' },
          { type: 'postback', label: '自営業・フリーランス', data: 'employment=自営業・フリーランス' },
          { type: 'postback', label: '派遣・契約社員',      data: 'employment=派遣・契約社員' }
        ]
      }
    }]);
    return;
  }

  // 雇用形態
  if (params.employment) {
    state.answers.employment = params.employment;
    state.step = 'ASK_SELFFUND';
    setState(userId, state);
    await reply(event.replyToken, [{
      type: 'template',
      altText: '自己資金を選んでください',
      template: {
        type: 'buttons',
        title: '頭金（自己資金）のご用意は？',
        text: '物件価格に対する自己資金の目安',
        actions: [
          { type: 'postback', label: 'なし（フルローン希望）', data: 'selfFund=なし（フルローン希望）' },
          { type: 'postback', label: '〜200万円',             data: 'selfFund=〜200万円' },
          { type: 'postback', label: '200〜500万円',          data: 'selfFund=200〜500万円' },
          { type: 'postback', label: '500万円以上',           data: 'selfFund=500万円以上' }
        ]
      }
    }]);
    return;
  }

  // 自己資金
  if (params.selfFund) {
    state.answers.selfFund = params.selfFund;
    state.step = 'ASK_OTHERLOAN';
    setState(userId, state);
    await reply(event.replyToken, [{
      type: 'template',
      altText: '他のローンを選んでください',
      template: {
        type: 'buttons',
        title: '現在、他にローンはありますか？',
        text: '車・カード・教育ローンなど',
        actions: [
          { type: 'postback', label: 'なし',          data: 'otherLoan=なし' },
          { type: 'postback', label: '月2万円未満',   data: 'otherLoan=月2万円未満' },
          { type: 'postback', label: '月2〜5万円',    data: 'otherLoan=月2〜5万円' },
          { type: 'postback', label: '月5万円以上',   data: 'otherLoan=月5万円以上' }
        ]
      }
    }]);
    return;
  }

  // 他ローン
  if (params.otherLoan) {
    state.answers.otherLoan = params.otherLoan;
    state.step = 'ASK_TIMING';
    setState(userId, state);
    await reply(event.replyToken, [{
      type: 'template',
      altText: '検討時期を選んでください',
      template: {
        type: 'buttons',
        title: 'ご購入の検討時期は？',
        text: 'いつ頃をお考えですか？',
        actions: [
          { type: 'postback', label: 'すぐにでも',        data: 'timing=すぐにでも' },
          { type: 'postback', label: '3ヶ月以内',         data: 'timing=3ヶ月以内' },
          { type: 'postback', label: '半年〜1年以内',     data: 'timing=半年〜1年以内' },
          { type: 'postback', label: 'まだ情報収集中',    data: 'timing=情報収集中' }
        ]
      }
    }]);
    return;
  }

  // 検討時期
  if (params.timing) {
    state.answers.timing = params.timing;
    state.step = 'ASK_NAME';
    setState(userId, state);
    await reply(event.replyToken, [{
      type: 'text',
      text: 'ありがとうございます！\nほぼ完了です。\n\n最後にお名前をお聞かせください。\n（ニックネームでもOKです）'
    }]);
    return;
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// テキストメッセージ処理
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async function handleMessage(event) {
  const userId = event.source.userId;
  const text   = event.message.text.trim();
  const state  = getState(userId);

  switch (state.step) {
    case 'ASK_AREA':
      state.answers.area = text;
      state.step = 'ASK_PROPERTY_TYPE';
      setState(userId, state);
      await reply(event.replyToken, [{
        type: 'template',
        altText: '物件種別を選んでください',
        template: {
          type: 'buttons',
          title: 'ご検討の物件種別は？',
          text: '当てはまるものをお選びください',
          actions: [
            { type: 'postback', label: '新築マンション',   data: 'propertyType=新築マンション' },
            { type: 'postback', label: '中古マンション',   data: 'propertyType=中古マンション' },
            { type: 'postback', label: '新築・中古戸建',   data: 'propertyType=戸建' },
            { type: 'postback', label: 'まだ決まってない', data: 'propertyType=未定' }
          ]
        }
      }]);
      break;

    case 'ASK_NAME':
      state.answers.name = text;
      await completeConsult(event, userId, state);
      break;

    default:
      await reply(event.replyToken, [{
        type: 'text',
        text: 'ご連絡ありがとうございます！\n「相談したい」と送っていただければ\n最初からご案内をスタートします。'
      }]);
      break;
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 相談完了処理
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async function completeConsult(event, userId, state) {
  const a         = state.answers;
  const timestamp = new Date().toLocaleString('ja-JP', { timeZone: 'Asia/Tokyo' });
  const sim       = generateSimulation(a);

  // スプレッドシートに記録
  await writeToSheet({
    timestamp, userId,
    name:           a.name || '',
    consultType:    a.consultType || '',
    area:           a.area || '',
    propertyType:   a.propertyType || '',
    price:          a.price || '',
    income:         a.income || '',
    employment:     a.employment || '',
    selfFund:       a.selfFund || '',
    otherLoan:      a.otherLoan || '',
    timing:         a.timing || '',
    loanAmount:     formatYen(sim.loanAmount),
    monthlyPayment: formatYen(sim.monthly_v1),
    ratio:          sim.ratio
  });

  console.log(`[新規相談] ${a.name} / ${a.consultType} / 返済比率${sim.ratio}%`);

  // お客さまへ完了メッセージ
  await reply(event.replyToken, [
    {
      type: 'text',
      text: `${a.name}様、ご回答ありがとうございます！`
    },
    {
      type: 'text',
      text: sim.message
    },
    {
      type: 'text',
      text: `📋 ご入力内容\n\n種別：${a.consultType || '−'}\nエリア：${a.area || '−'}\n物件：${a.propertyType || '−'}\n価格：${a.price || '−'}\n年収：${a.income || '−'}\n雇用：${a.employment || '−'}\n自己資金：${a.selfFund || '−'}\n他ローン：${a.otherLoan || '−'}\n検討時期：${a.timing || '−'}`
    },
    {
      type: 'template',
      altText: '次のステップ',
      template: {
        type: 'buttons',
        title: '次のステップ',
        text: '担当者に詳しく相談しますか？',
        actions: [
          { type: 'postback', label: '📞 担当者に相談したい', data: 'next=consult' },
          { type: 'postback', label: '🏠 物件も一緒に探したい', data: 'next=property' },
          { type: 'postback', label: '📩 後日連絡を希望', data: 'next=later' }
        ]
      }
    }
  ]);

  // ヒアリング完了後は自動で手動モードに切り替え
  setManual(userId);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// LINE APIへ返信
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async function reply(replyToken, messages) {
  try {
    const res = await fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`
      },
      body: JSON.stringify({ replyToken, messages })
    });
    if (!res.ok) console.error('LINE API Error:', await res.text());
  } catch (err) {
    console.error('Reply error:', err);
  }
}

function parseData(str) {
  const r = {};
  str.split('&').forEach(p => {
    const [k, v] = p.split('=');
    if (k && v) r[decodeURIComponent(k)] = decodeURIComponent(v);
  });
  return r;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// サーバー起動
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
app.listen(PORT, () => {
  console.log(`住宅ローン相談Bot running on port ${PORT}`);
  console.log('担当者コマンド: #対応開始 / #bot再開 / #状態確認');
});
