// === 定数定義 ===
/** OpenAI APIキー */
const OPENAI_API_KEY = 'YOUR_API_KEY';

/** OpenAIモデル */
const OPENAI_MODEL = 'gpt-4o';

/** 最大トークン数 */
const MAX_TOKENS = 2000;

/** 応答の創造性設定 */
const TEMPERATURE = 0.7;

/** ユーザー名 */
const USER_NAME = 'YOUR_NAME';

/** 組織名 */
const ORGANIZATION = 'YOUR_ORGANIZATION';

/** ユーザー肩書き */
const USER_TITLE = 'YOUR_TITLE';

/** ユーザーのメールアドレス */
const USER_EMAIL = 'YOUR_EMAIL';

/** メール署名 */
const MAIL_SIGNATURE = `
--${ORGANIZATION}
${USER_NAME}
${USER_TITLE}
email: ${USER_EMAIL}
`;

/** アクションリスト */
const ACTIONS_LIST = [
  '賛成',
  '感謝',
  '反対',
  '新しい提案',
  '同意',
  '承諾するけれど確定まで待ってほしい',
  'その日時だと都合が悪いので別の日時を提案してほしい',
];

/**
 * Gmailアドオンのメイン関数: スレッド内容とアクションボタンを表示します。
 * @param {!Object} event Gmailアドオンイベントオブジェクト。
 * @return {!CardService.Card} GmailアドオンのカードUI。
 */
function getContextualAddOn(event) {
  const accessToken = event.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);

  const threadId = event.messageMetadata.threadId;
  const thread = GmailApp.getThreadById(threadId);
  const messages = thread.getMessages();

  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('スレッド内容とアクション'));

  // スレッド内のメッセージを表示
  messages.forEach((message, index) => {
    const section = CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText(
            `<b>メッセージ ${index + 1}</b><br>` +
            `送信者: ${message.getFrom()}<br>` +
            `件名: ${message.getSubject()}<br>` +
            `本文: ${message.getPlainBody().slice(0, 500)}...`
        )
    );
    card.addSection(section);
  });

  // アクションボタンを表示するセクション
  const actionSection = CardService.newCardSection();

  ACTIONS_LIST.forEach((action) => {
    actionSection.addWidget(
        CardService.newTextButton()
            .setText(action)
            .setOnClickAction(
                CardService.newAction()
                    .setFunctionName('generateReplyDraft')
                    .setParameters({action})
            )
    );
  });

  card.addSection(actionSection);
  return card.build();
}

/**
 * ChatGPTを使用して返信下書きを生成し、スレッドに保存します。
 * @param {!Object} event Gmailアドオンイベントオブジェクト。
 * @return {!CardService.Card} 下書き保存結果を表示するカードUI。
 */
function generateReplyDraft(event) {
  const userAction = event.parameters.action;
  const threadId = event.messageMetadata.threadId;
  const thread = GmailApp.getThreadById(threadId);
  const messages = thread.getMessages();

  const latestMessage = messages[messages.length - 1];
  const context = latestMessage.getPlainBody();
  const subject = latestMessage.getSubject();

  // ChatGPT APIにリクエストを送信して返信を生成
  const replyDraft = callChatGPTAPI(userAction, context);

  // 返信下書きをスレッドに紐づけて保存
  const draftId = saveDraft(replyDraft, latestMessage.getFrom(), subject, threadId);

  // 結果を表示
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('返信下書きがスレッドに保存されました'));

  card.addSection(
      CardService.newCardSection().addWidget(
          CardService.newTextParagraph().setText(
              `<b>${userAction}</b> に基づく返信の下書きがスレッドに紐づけられて保存されました。ドラフトID: ${draftId}`
          )
      )
  );

  return card.build();
}

/**
 * Gmailスレッドに返信下書きを保存します。
 * @param {string} replyContent 返信内容。
 * @param {string} recipient 返信先のメールアドレス。
 * @param {string} subject メール件名。
 * @param {string} threadId GmailスレッドID。
 * @return {string} 作成されたドラフトのID。
 */
function saveDraft(replyContent, recipient, subject, threadId) {
  const thread = GmailApp.getThreadById(threadId);
  const draft = thread.createDraftReply(replyContent);
  return draft.getId();
}

/**
 * ChatGPT APIを呼び出して返信を生成します。
 * @param {string} action ユーザーが選択したアクション。
 * @param {string} context メール本文。
 * @return {string} 生成された返信内容。
 */
function callChatGPTAPI(action, context) {
  const apiUrl = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: OPENAI_MODEL,
    messages: [
      {
        role: 'system',
        content: `${USER_NAME}です。ビジネスメールの返信を書いています。`,
      },
      {
        role: 'user',
        content: `以下の内容のビジネスメールに対して、${USER_NAME}として返信を書いています。${action}の姿勢で丁寧な返信の下書きを作成してください。署名は${MAIL_SIGNATURE}でお願いします。\n\n${context}`,
      },
    ],
    max_tokens: MAX_TOKENS,
    temperature: TEMPERATURE,
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(apiUrl, options);
  const json = JSON.parse(response.getContentText());

  if (json.choices && json.choices.length > 0) {
    return json.choices[0].message.content;
  } else {
    return '返信下書きの生成に失敗しました。';
  }
}

