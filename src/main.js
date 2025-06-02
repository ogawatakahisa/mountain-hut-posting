// main.js
function getMailAndWrite(label) {
  let query;
  let today = new Date();
  let oneDayAgo = new Date(today.getTime() - (10 * 60 * 60 * 1000)); // 1日前の日時
  let oneDayAgoFormatted = Utilities.formatDate(oneDayAgo, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  if (label === "usuall") {
      // 受信トレイの中で、件名が「ご予約依頼を受付いたしました」となっており、特定のラベルやRe、FWが含まれていないメールを検索
      query = `in:inbox subject:ご予約依頼を受付いたしました -(label:Success OR label:MailCh OR label:SheetCh OR label:想定外エラー OR label:未作成のブック OR label:not氏名in予約表 OR subject:"Re: " OR subject:"FW: ")`;

  } else if (label === "予約上限") {
      // 件名に「Re」が含まれず、「予約上限」ラベルに該当する件名
      query = 'label:予約上限 subject:ご予約依頼を受付いたしました -(subject:"Re: " OR subject:"FW: ")';

  } else if (label === "未作成のブック") {
      query = 'label:未作成のブック subject:ご予約依頼を受付いたしました -(subject:"Re: " OR subject:"FW: ")';

  } else if (label === "未作成のシート") {
      query = `label:未作成のシート subject:ご予約依頼を受付いたしました -(subject:"Re: " OR subject:"FW: ")`;

  } else if (label === "not氏名") {
      query = 'label:not氏名in予約表 subject:ご予約依頼を受付いたしました -(subject:"Re: " OR subject:"FW: ")';

  } else if (label === "想定外エラー") {
      // 件名に「Re」が含まれず、「想定外のエラー」ラベルに該当する件名
      query = `label:想定外エラー subject:ご予約依頼を受付いたしました -(subject:"Re: " OR subject:"FW: ")`;

  } else if (label === "Test") {
      query = 'label:Test -subject:"Re: "';
  }


  Logger.log("メール取得件数: " + query);
  let gmail = getGmail(query);
  Logger.log("メール取得件数: " + gmail.length);

    if (gmail.length > 0) {

      let rowLimit = 108 // 全山荘共通で予約上限は100件
      let fgetmail = true // メールから想定通りに情報を取得できたか
      
      gmail.forEach(email => {
          Logger.log("取得した件名: " + email.subject);
          // Logger.log("実行したメールを検索する"+'{"' + email.applicantName + '" AND "' + email.phoneNumberInfo + '"}')
          
          let allSuccessful = true; // 全ての日程（連泊）の転記の成功
          let writtenData = []; // 転記したデータを追跡する配列
          fgetmail = true

          try {
              let nights = email.nights; // 泊数を取得

              // 件名から基準日を取得
              let baseDate = getDateFromSubject(email.subject);

              if (!baseDate) {
                addLabelToEmail(email.threadId, "イレギュラー日付");
                allSuccessful = false;
                return;
              }

              let stayLabel = ""; // 連泊
              let fgetmail = true //メールから想定通りに取得できたか

              // 一日分ずつ転記する
              for (let i = 0; i < nights; i++) {
                Utilities.sleep(1500);
                  
                  let targetDate = new Date(baseDate);
                  // 月、年跨ぎ
                  targetDate.setDate(targetDate.getDate() + i);

                  // ブック名
                  let year = targetDate.getFullYear();
                  let month = targetDate.getMonth() + 1;
                  let lodgeName = email.lodgeName;
                  
                  let bookName = `${lodgeName}_${year}-${month}月`;

                  if (year === 2025) {
                    if (month <= 3) {
                      addLabelToEmail(email.threadId, "テスト範囲外");
                      allSuccessful = false;
                      break;
                    }
                  }

                  let folderPas = ""
                  if (lodgeName === "夏沢鉱泉") {
                    folderPas = "00_各山荘予約表/30_夏沢予約表"
                    // folderPas = "00_各山荘予約表/30_夏沢予約表/夏沢鉱泉_2025-4月～2026-3月/テスト中(事務所用)"

                  } else if (lodgeName === "根石岳山荘") {
                    folderPas = "00_各山荘予約表/20_根石予約表"
                    // folderPas = "00_各山荘予約表/20_根石予約表/根石岳山荘_2025-4月～2026-3月/テスト中(事務所用)"

                  } else if (lodgeName === "硫黄岳山荘") {
                    folderPas = "00_各山荘予約表/10_硫黄予約表"
                    // folderPas = "00_各山荘予約表/10_硫黄予約表/硫黄岳山荘_2025-4月～2026年-3月/テスト中(事務所用)"
                  }

                  let book = getSpreadsheetByName(folderPas , bookName);

                  if (!book) {
                      addLabelToEmail(email.threadId, "未作成のブック");
                      allSuccessful = false;
                      break;
                  }


                  // シート名
                  let sheetName = getSheetNameFromDate(targetDate);
                  let sheet = book.getSheetByName(sheetName);
                  if (!sheet) {
                      addLabelToEmail(email.threadId, "未作成のシート");
                      allSuccessful = false;
                      break;
                  }
                
                  // msgID重複チェック
                  const existingIdsRange = sheet.getRange(8, 41, 101);
                  const existingIds      = existingIdsRange.getValues().flat(); 

                  const thisMsgId = email.msgId;
                  if (existingIds.includes(thisMsgId)) {
                    Logger.log(`重複スキップ：メール msgId=${thisMsgId} は既に AO 列に存在`);
                    return;
                  }
                
                  // 氏名が空の行を転記先とする
                  let receptionColumn = findColumn(sheet, "氏名");

                  // 氏名が見つからない場合は転記不可のため
                  if (receptionColumn === -1) {
                      addLabelToEmail(email.threadId, "not氏名in予約表");
                      allSuccessful = false;
                      break;
                  }
                  // 転記する行を探す
                  let lastRow = findFirstEmptyRow(sheet, receptionColumn);

                  // 予約上限に達しているかの確認
                  if (lastRow >= rowLimit) {
                      addLabelToEmail(email.threadId, "予約上限");
                      allSuccessful = false;
                      break;
                  }

                  let pickupTime = email.pickupTime;
                  let dropoffTime = email.dropoffTime;

                  // let stayLabel; // 連泊
                  // 泊数に応じた連泊の値を設定
                  if (nights >= 2) {
                      stayLabel = i === 0 ? "初日" : `${convertToFullWidthNumber(i + 1)}日目`;

                      // 連泊の場合送迎情報を適切に分ける処理
                      pickupTime = i === 0 ? email.pickupTime : "なし"; // 初日だけ「迎え」
                      dropoffTime = i === nights - 1 ? email.dropoffTime : "なし"; // 最終日だけ「送り」
                      if (i === nights - 1) {
                          dropoffTime = email.dropoffTime;
                      } else {
                          dropoffTime = "なし";
                      }
                  }

                  // 連絡事項がある場合は備考欄にメール検索用の文字列を書く
                  let serchmail = "";
                  if (email.contact) {
                      serchmail = '{"' + email.applicantName + '" AND "' + email.phoneNumberInfo + '"}';


                  }
                  // 転記先と内容
                  let data = [
                      { column: "受付日", value: email.formattedDate },
                      { column: "受付方法", value: email.mail },
                      { column: "氏名", value: email.applicantName },
                      { column: "都道府県", value: email.prefecture },
                      { column: "人数", value: email.totalGuests },
                      { column: "男", value: email.maleGuests },
                      { column: "女", value: email.femaleGuests },
                      { column: "中", value: email.juniorHighsCount },
                      { column: "小", value: email.elementarysCount },
                      { column: "幼", value: email.toddlersCount },
                      { column: "食事", value: email.meal },
                      { column: "個室", value: email.roomSharingInfo },
                      { column: "送迎場所", value: email.pickupPlace },
                      { column: "迎え時間", value: pickupTime },
                      { column: "送り時間", value: dropoffTime },
                      { column: "電話番号", value: email.phoneNumber },
                      { column: "登山口", value: email.startingPoint },
                      { column: "状況", value: "未確認" },
                      { column: "連泊", value: stayLabel },
                      { column: "備考", value: serchmail }
                  ];

                  for (let entry of data) {
                      // 転記先の項目が存在するか
                      entry.columnIndex = findColumn(sheet, entry.column);

                      if (entry.columnIndex === -1) {
                          // 転記先の項目が見つからない
                          // addLabelToEmail(email.threadId, "SheetCh");
                          // 「備考」に serchmail を追加で記録
                          let remarksColumnIndex = findColumn(sheet, "備考");
                          if (remarksColumnIndex !== -1) {
                              let remarksRange = sheet.getRange(lastRow, remarksColumnIndex + 1);
                              let currentValue = remarksRange.getValue();
                              remarksRange.setValue(currentValue + "\n" + serchmail);
                              addLabelToEmail(email.threadId, "SheetCh");
                          } else {
                              addLabelToEmail(email.threadId, "not備考in予約表");
                          }

                      } else {
                          Logger.log("項目名"+entry.column)
                          Logger.log("値"+entry.value)
                          Logger.log("項目場所"+entry.columnIndex)
                          // 転記先を定めて転記
                          let range = sheet.getRange(lastRow, entry.columnIndex + 1);
                          range.setValue(entry.value);
                          if (entry.value.includes("エラー")) {
                            fgetmail = false
                          }

                          // 転記したセルを記録
                          //  要は何行目の何列目まで転記したのかを記録しておいて、もし、エラーになったらここに記録されている、列、行をもとに削除する 
                          writtenData.push({ sheet: sheet, row: lastRow, column: entry.columnIndex + 1 });
                      }
                  }
                  sheet.getRange(lastRow, 41).setValue(thisMsgId);
              }

          // 連泊で転記できないエラーがあった時のための対策をしていたがロジックを変更する必要がある
          // ↓
          // 転記できない===赤色のラベルが付くとき
          // if (!allSuccessful) {
              // for (let entry of writtenData) {
              //     let cell = entry.sheet.getRange(entry.row, entry.column);
              //     cell.clearContent();
              // }
          // }

          } catch (error) {
              Logger.log("メール処理中にエラーが発生しました (件名: " + email.subject + "): " + error.message);
              if (error.message.includes("Service Spreadsheets failed") || error.message.includes("スプレッドシートのサービスに接続できなくなりました")) {
                addLabelToEmail(email.threadId, "想定外エラー/接続エラー");
              } else {
                addLabelToEmail(email.threadId, "想定外エラー");
              }
              allSuccessful = false;

              // 転記したデータを削除
              for (let entry of writtenData) {
                  let cell = entry.sheet.getRange(entry.row, entry.column);
                  cell.clearContent();
              }
          }

          // 全泊数分の転記が成功した場合のみラベルを付与
          if (allSuccessful) {
              addLabelToEmail(email.threadId, "Success");

              if (fgetmail === false) {
                addLabelToEmail(email.threadId, "MailCh");
              }
          }
      });
  }

}
function onButtonClickUsuall() {
  getMailAndWrite('usuall');
}
function onButtonClickLimit() {
  getMailAndWrite('予約上限');
}
function onButtonClickNotyet() {
  getMailAndWrite('未作成のブック');
}
function onButtonClickError() {
  getMailAndWrite('想定外エラー');
}
function onButtonClickNotyetSheet() {
  getMailAndWrite('未作成のシート');
}
function onButtonClicknotName() {
  getMailAndWrite('not氏名');
}
// 他のボタンも同様
