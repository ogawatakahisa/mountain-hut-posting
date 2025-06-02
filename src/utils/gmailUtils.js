// 全く同じ件名があってもReとか省く。
// 全く同じ件名でも、メールの内容が異なる場合は別のメールとして扱う。
// let message = messages[0];はreturnがあっても転記しないためのもの
// →でもこれがあることで全く同じ件名の異なる内容のメールが転記されない状態になってる

function getGmail(query) {
  let threads = GmailApp.search(query);
  Logger.log("日付: " + threads);

  // 古い順に処理
  threads.reverse();

  let gmailInfo = [];

  threads.forEach(thread => {
    // スレッド内の全メッセージを処理
    let messages = thread.getMessages();

    messages.forEach(message => {
      let subject = message.getSubject();

      // Re:／FW: はスキップ
      if (subject.startsWith("Re:") || subject.startsWith("FW:")) return;

      // 差出人チェック
      if (!message.getFrom().includes("wordpress@iodake.jp")) return;

      try {
        const msgId = message.getId();
        Logger.log("メールID   : " + msgId);
        Logger.log("件名     : " + subject);

        let plainBody = message.getPlainBody();
        let receivedTime = message.getDate();
        let formattedDate = Utilities.formatDate(receivedTime, Session.getScriptTimeZone(), 'yyyy/MM/dd');
        let lodgeName = extractLodgeName(subject);
        let nightsMatch = plainBody.match(/泊数：(\d+)/);
        let nights = nightsMatch ? parseInt(nightsMatch[1], 10) : 1;
        let guestsInfo = extractGuestsInfo(plainBody);
        let phoneNumberInfo = formatPhoneNumber(plainBody);
        let formattedPhoneNumber = phoneNumberInfo[0];
        let origPhoneNumber = phoneNumberInfo[1];

        gmailInfo.push({
          subject: subject,
          formattedDate: formattedDate,
          lodgeName: lodgeName,
          nights: nights,
          mail: "メール",
          applicantName: (plainBody.match(/申込者氏名：(.*)/) || ["", "エラー"])[1].trim(),
          prefecture: extractPrefecture(plainBody),
          totalGuests: guestsInfo.totalGuests,
          maleGuests: guestsInfo.maleGuests,
          femaleGuests: guestsInfo.femaleGuests,
          juniorHighsCount: guestsInfo.juniorHighsCount,
          elementarysCount: guestsInfo.elementarysCount,
          toddlersCount: guestsInfo.toddlersCount,
          meal: formatMealDescription((plainBody.match(/食事：([^\n]+)/) || ["", "エラー"])[1]),
          roomSharingInfo: getRoomInfo(plainBody, lodgeName),
          pickupPlace: getPickupPlaceInfo(plainBody),
          pickupTime: getPickupTimeInfo(plainBody, "pickup"),
          dropoffTime: getPickupTimeInfo(plainBody, "dropoff"),
          phoneNumber: formattedPhoneNumber,
          startingPoint: (plainBody.match(/入山口：(.*)/) || ["", "エラー"])[1].trim(),
          contact: getContactInfo(plainBody),
          threadId: thread.getId(),
          plainBody: plainBody,
          phoneNumberInfo: origPhoneNumber
        });

      } catch (err) {
        console.log("エラーが発生しました：スレッド件名 = " + thread.getFirstMessageSubject());
      }
    });
  });

  return gmailInfo;
}


function addLabelToEmail(threadId, labelName) {
  let fullPathLabelName = labelName;
  // ネストされたラベルのフルパスを設定
  if (labelName === LABELS.MAIL_CH || labelName === LABELS.SHEET_CH || labelName === "not備考in予約表") {
    fullPathLabelName = LABELS.SUCCESS + "/" + labelName;
  }

  let label = GmailApp.getUserLabelByName(fullPathLabelName) || GmailApp.createLabel(fullPathLabelName);
  let thread = GmailApp.getThreadById(threadId);
  thread.addLabel(label);

  // 不要なラベルを削除
  if (labelName === LABELS.SUCCESS) {
    const labelsToRemove = [
      LABELS.ERROR,
      LABELS.NOT_YET_BOOK,
      LABELS.NOT_YET_SHEET,
      LABELS.RESERVATION_LIMIT,
      LABELS.REJECTED_DATE,
      LABELS.NOT_NAME
    ];
    labelsToRemove.forEach(labelToRemove => {
      removeLabelFromEmail(threadId, labelToRemove);
    });

  } else if (labelName === LABELS.NOT_YET_SHEET) {
    removeLabelFromEmail(threadId, LABELS.NOT_YET_BOOK);

  } else if (labelName === LABELS.RESERVATION_LIMIT) {
    removeLabelFromEmail(threadId, LABELS.NOT_YET_BOOK);
    removeLabelFromEmail(threadId, LABELS.NOT_YET_SHEET);

  } else if (labelName === LABELS.NOT_NAME) {
    removeLabelFromEmail(threadId, LABELS.RESERVATION_LIMIT);
    removeLabelFromEmail(threadId, LABELS.NOT_YET_BOOK);
    removeLabelFromEmail(threadId, LABELS.NOT_YET_SHEET);

  } else if (labelName === LABELS.TEST_RANGE) {
    removeLabelFromEmail(threadId, LABELS.REJECTED_DATE);
  }
}

function removeLabelFromEmail(threadId, labelName) {
  let fullPathLabelName = labelName;
  if (labelName === LABELS.MAIL_CH || labelName === LABELS.SHEET_CH || labelName === "not備考in予約表") {
    fullPathLabelName = LABELS.SUCCESS + "/" + labelName;
  }

  let label = GmailApp.getUserLabelByName(fullPathLabelName);
  if (label) {
    let thread = GmailApp.getThreadById(threadId);
    thread.removeLabel(label);
  }
}
