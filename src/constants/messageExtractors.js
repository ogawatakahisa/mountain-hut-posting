// messageExtractors.js

// 既存の関数内容の全てを書き出します。

function extractLodgeName(subject) {
  for (const name of LODGE_NAMES) {
    if (subject.includes(name)) return name;
  }
  return "エラー";
}

function extractNights(plainBody) {
  const match = plainBody.match(/泊数：(\d+)/);
  return match ? parseInt(match[1], 10) : 1;
}

// 他の抽出関数も同様に
// 例：extractPrefecture, getRoomInfo, extractGuestsInfo, etc.
