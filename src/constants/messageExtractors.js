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