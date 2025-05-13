// gmailUtils.js
const { LABELS } = constants;

function getGmail(query) {
  const threads = GmailApp.search(query);
  threads.reverse();
  const gmailInfo = [];
  threads.forEach(thread => {
    const messages = thread.getMessages();
    const message = messages[0];
    const subject = message.getSubject();
    if (subject.startsWith("Re:") || subject.startsWith("FW:")) return;
    if (!message.getFrom().includes("wordpress@iodake.jp")) return;

    const plainBody = message.getPlainBody();
    const receivedTime = message.getDate();
    const formattedDate = Utilities.formatDate(receivedTime, Session.getScriptTimeZone(), 'yyyy/MM/dd');

    const emailData = {
      subject: subject,
      formattedDate: formattedDate,
      lodgeName: messageExtractors.extractLodgeName(subject),
      nights: messageExtractors.extractNights(plainBody),
      mail: "メール",
      applicantName: (plainBody.match(/申込者氏名：(.*)/) || ["", "エラー"])[1].trim(),
      // 他の抽出関数も呼び出し
      // ...
      threadId: thread.getId(),
      plainBody: plainBody,
      phoneNumberInfo: formatUtils.formatPhoneNumber(plainBody)[1]
    };
    gmailInfo.push(emailData);
  });
  return gmailInfo;
}

function addLabelToEmail(threadId, labelName) {
  let fullPathLabelName = labelName;
  if (labelName === LABELS.MAIL_CH || labelName === LABELS.NOT_NAME || labelName === LABELS.SHEET_CH) {
    fullPathLabelName = "Success/" + labelName;
  }
  let label = GmailApp.getUserLabelByName(fullPathLabelName) || GmailApp.createLabel(fullPathLabelName);
  let thread = GmailApp.getThreadById(threadId);
  thread.addLabel(label);
  // ラベル削除処理も同様に
  // ...
}

function removeLabelFromEmail(threadId, labelName) {
  // 同様に
}
