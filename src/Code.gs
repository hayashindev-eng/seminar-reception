// ─────────────────────────────────────────────────
//  エントリポイント & ユーティリティ
// ─────────────────────────────────────────────────

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'reception';

  let template;
  if (page === 'settings') {
    template = HtmlService.createTemplateFromFile('settings');
  } else {
    template = HtmlService.createTemplateFromFile('index');
  }

  return template.evaluate()
    .setTitle('セミナー受付システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// HTML ファイルのインクルード用ヘルパー
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
