
// 電話番号にハイフンをつける関数
function formatPhoneNumber(plainBody) {
  let orignumber;
  let number;

  orignumber = (plainBody.match(/電話番号：(.*)/) || ["", "エラー"])[1].trim();
  number = orignumber.replace(/[-\s]/g, '');

  // 携帯電話 (例: 080-1234-5678)
  if (/^(090|080|070)\d{8}$/.test(number)) {
    let formattedNumber = number.slice(0, 3) + '-' + number.slice(3, 7) + '-' + number.slice(7);
    return [formattedNumber, orignumber];
  }

  // 固定電話 (市外局番が1桁～4桁)
  if (/^0\d{9,10}$/.test(number)) {
    // 市外局番が1桁 (例: 03-XXXX-XXXX)
    if (/^0[1-9]\d{8}$/.test(number)) {
      let formattedNumber = number.slice(0, 2) + '-' + number.slice(2, 6) + '-' + number.slice(6);
      return [formattedNumber, orignumber];
    }
    // 市外局番が2桁 (例: 06-XXXX-XXXX)
    else if (/^0[1-9]{2}\d{8}$/.test(number)) {
      let formattedNumber = number.slice(0, 3) + '-' + number.slice(3, 7) + '-' + number.slice(7);
      return [formattedNumber, orignumber];
    }
    // 市外局番が3桁 (例: 045-XXX-XXXX)
    else if (/^0[1-9]{3}\d{7}$/.test(number)) {
      let formattedNumber = number.slice(0, 4) + '-' + number.slice(4, 7) + '-' + number.slice(7);
      return [formattedNumber, orignumber];
    }
    // 市外局番が4桁 (例: 0564-XX-XXXX)
    else if (/^0[1-9]{4}\d{6}$/.test(number)) {
      let formattedNumber = number.slice(0, 4) + '-' + number.slice(4, 6) + '-' + number.slice(6);
      return [formattedNumber, orignumber];
    }
  }

  // 不正な電話番号の場合はそのまま返す
  return [number, orignumber];
}

// シート名を取得する
function getSheetNameFromDate(date) {
  let day = date.getDate();
  let dayOfWeek = DAYS_OF_WEEK[date.getDay()];
  return `${day}（${dayOfWeek}）`;
}

// 件名から基準日を取得
function getDateFromSubject(subject) {
  let datePattern = /(\d{4})\/(\d{2})\/(\d{2})/; //YYYY/MM/DDにのみ対応
  let matches = subject.match(datePattern);
  if (matches) {
    let year = parseInt(matches[1], 10);
    let month = parseInt(matches[2], 10) - 1;
    let day = parseInt(matches[3], 10);
    return new Date(year, month, day);
  }
  return null;
}

// 部屋情報の抽出と変換
function getRoomInfo(plainBody, lodgeName) {
  let prevalue;
  if (lodgeName === "夏沢鉱泉") {
    prevalue = (plainBody.match(/相部屋可否：\s*([^\n]*)/) || ["", "エラー"]);
  } else if (lodgeName === "硫黄岳山荘" || lodgeName === "根石岳山荘") {
    prevalue = (plainBody.match(/個室希望有無：\s*([^\n]*)/) || ["", "エラー"]);
  } else {
    return "山荘名が不正";
  }

  let roomSharingInfo = "";
  if (prevalue && prevalue[1]) {
    roomSharingInfo = prevalue[1].trim();
  }

  // 山荘名に応じた部屋情報の変換
  if (lodgeName === "夏沢鉱泉") {
    if (roomSharingInfo.includes("男女別相部屋可")) {
      roomSharingInfo = "別相可";
    } else if (roomSharingInfo.includes("男女相部屋可")) {
      roomSharingInfo = "相可";
    } else if (roomSharingInfo.includes("相部屋不可")) {
      roomSharingInfo = "相不可";
    }
  } else if (lodgeName === "根石岳山荘") {
    if (roomSharingInfo.includes("個室希望なし")) {
      roomSharingInfo = "";
    } else if (roomSharingInfo.includes("個室希望")) {
      roomSharingInfo = "希望";
    } else if (roomSharingInfo.includes("個室条件")) {
      roomSharingInfo = "条件";
    }
  } else if (lodgeName === "硫黄岳山荘") {
    if (roomSharingInfo.includes("個室希望なし")) {
      roomSharingInfo = "";
    } else if (roomSharingInfo.includes("一般個室希望")) {
      roomSharingInfo = "一般希望";
    } else if (roomSharingInfo.includes("ラグジュアリー個室希望")) {
      roomSharingInfo = "ラグ希望";
    } else if (roomSharingInfo.includes("個室条件")) {
      roomSharingInfo = "条件";
    }
  }
  return roomSharingInfo;
}

// 人数情報の抽出
function extractGuestsInfo(plainBody) {
  // 合計人数計算用の変数
  let sumresult = 0;
  let inputjuniorhigh;
  let inputerement;
  let inputTodd;

  // 中学生の人数の取得
  let menjunior = (plainBody.match(/中学生（男の子）：(\d+)/) || ["", "0"])[1];
  let womjunior = (plainBody.match(/中学生（男の子）：\s*\d+人\s*｜\s*（女の子）：\s*(\d+)人/) || ["", "0"])[1];

  let juniorboy = parseInt(menjunior, 10);
  let juniorgirl = parseInt(womjunior, 10);
  sumresult = juniorboy + juniorgirl;
  if (sumresult >= 1) {
    inputjuniorhigh = sumresult;
  } else {
    inputjuniorhigh = "";
  }

  // 小学生の人数の取得
  let menerement = (plainBody.match(/小学生（男の子）：(\d+)/) || ["", "0"])[1];
  let womerement = (plainBody.match(/小学生（男の子）：\s*\d+人\s*｜\s*（女の子）：\s*(\d+)人/) || ["", "0"])[1];

  let erementboy = parseInt(menerement, 10);
  let erementirl = parseInt(womerement, 10);
  sumresult = erementboy + erementirl;
  if (sumresult >= 1) {
    inputerement = sumresult;
  } else {
    inputerement = "";
  }

  // 幼児の人数の取得
  let menkids = (plainBody.match(/幼児（男の子）：(\d+)/) || ["", "0"])[1];
  let womkids = (plainBody.match(/幼児（男の子）：\s*\d+人\s*｜\s*（女の子）：\s*(\d+)人/) || ["", "0"])[1];

  let chboy = parseInt(menkids, 10);
  let chgirl = parseInt(womkids, 10);
  sumresult = chboy + chgirl;
  if (sumresult >= 1) {
    inputTodd = sumresult;
  } else {
    inputTodd = "";
  }

  // 大人の人数
  let men = (plainBody.match(/大人（男性）：\s*(\d+)人\s*｜/) || ["", "0"])[1];
  let women = (plainBody.match(/（女性）：\s*(\d+)人/) || ["", "0"])[1];

  let chmen = parseInt(men, 10);
  let chwomen = parseInt(women, 10);

  // 性別ごとの人数の合計
  chmen = chmen + juniorboy + erementboy + chboy;
  chwomen = chwomen + juniorgirl + erementirl + chgirl;

  return {
    totalGuests: (plainBody.match(/宿泊者合計人数：(\d+)人/) || ["", "エラー"])[1],
    maleGuests: String(chmen),
    femaleGuests: String(chwomen),
    juniorHighsCount: String(inputjuniorhigh),
    elementarysCount: String(inputerement),
    toddlersCount: String(inputTodd)
  };
}

// 食事情報のフォーマット調整
function formatMealDescription(mealDescription) {
  let formatted = mealDescription.trim()
    .replace(/3食（夕食・朝食・弁当）/g, '２食+弁当')
    .replace(/2食（夕食・朝食）/g, '２食')
    .replace(/夕食\+弁当/g, '夕食+弁当')
    .replace(/朝食\+弁当/g, '朝食+弁当')
    .replace(/夕食のみ/g, '夕食')
    .replace(/朝食のみ/g, '朝食');
  return formatted;
}

// 送迎時間を取得する関数
function getPickupTimeInfo(plainBody, type) {
  let pattern = type === "pickup"
    ? /お迎え\s*(.+?)(?:\s*｜|$)/
    : /お送り\s*(.+?便)/;

  let match = plainBody.match(pattern);
  let time = "なし";

  if (match && match[1]) {
    time = match[1].trim();
    time = time === "不要" ? "なし" : time;

    if (type === "pickup") {
      let timeMatch = time.match(/(\d{1,2}):(\d{2})/);
      if (timeMatch) {
        let hour = timeMatch[1].padStart(2, '0');
        let minute = timeMatch[2];
        time = hour + minute;
      }
    } else {
      if (time.includes("午前")) {
        time = "午前";
      } else if (time.includes("午後")) {
        time = "午後";
      }
    }
  }
  return time;
}

// 連絡事項の有無を確認する関数
function getContactInfo(plainBody) {
  let contact = false;
  let precontact = plainBody.match(/連絡事項：\s*([\s\S]*)/);
  if (precontact) {
    let lines = precontact[1].split("\n");
    let nextLine = lines[0].trim();
    if (nextLine !== "--") {
      contact = true;
    }
  }
  return contact;
}

// 送迎場所の情報を取得する関数
function getPickupPlaceInfo(plainBody) {
  let pickupTime = getPickupTimeInfo(plainBody, "pickup");
  let dropoffTime = getPickupTimeInfo(plainBody, "dropoff");
  let place = "エラー";
  if (pickupTime === "なし" && dropoffTime === "なし") {
    place = "なし";
  } else if (pickupTime === "なし" && dropoffTime !== "なし") {
    place = "駅東口";
  } else {
    let pickupMatch = plainBody.match(/お迎え\s.*【(?:[^】]*\s)?(?:\d{1,2}:\d{2}\s*)?(.+?)】\s*｜/);
    if (pickupMatch && pickupMatch[1]) {
      place = pickupMatch[1].trim();
      if (place === "不要") {
        place = "なし";
      } else if (place === "茅野駅東口") {
        place = "駅東口";
      } else if (place === "分岐") {
        place = "分岐";
      } else if (place === "桜平駐車場(下)") {
        place = "桜平下Ｐ";
      } else {
        place = "エラー";
      }
    }
  }
  return place;
}

// 電話番号にハイフンをつける関数
function formatPhoneNumber(plainBody) {
  let orignumber;
  let number;

  orignumber = (plainBody.match(/電話番号：(.*)/) || ["", "エラー"])[1].trim();
  number = orignumber.replace(/[-\s]/g, '');

  if (/^(090|080|070)\d{8}$/.test(number)) {
    let formattedNumber = number.slice(0, 3) + '-' + number.slice(3, 7) + '-' + number.slice(7);
    return [formattedNumber, orignumber];
  }

  if (/^0\d{9,10}$/.test(number)) {
    if (/^0[1-9]\d{8}$/.test(number)) {
      let formattedNumber = number.slice(0, 2) + '-' + number.slice(2, 6) + '-' + number.slice(6);
      return [formattedNumber, orignumber];
    } else if (/^0[1-9]{2}\d{8}$/.test(number)) {
      let formattedNumber = number.slice(0, 3) + '-' + number.slice(3, 7) + '-' + number.slice(7);
      return [formattedNumber, orignumber];
    } else if (/^0[1-9]{3}\d{7}$/.test(number)) {
      let formattedNumber = number.slice(0, 4) + '-' + number.slice(4, 7) + '-' + number.slice(7);
      return [formattedNumber, orignumber];
    } else if (/^0[1-9]{4}\d{6}$/.test(number)) {
      let formattedNumber = number.slice(0, 4) + '-' + number.slice(4, 6) + '-' + number.slice(6);
      return [formattedNumber, orignumber];
    }
  }

  return [number, orignumber];
}
