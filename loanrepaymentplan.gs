function calculateAge(birthdate, currentDate) {
  var birthDate = new Date(birthdate);
  var current = new Date(currentDate);

  // 年齢の基本計算
  var age = current.getFullYear() - birthDate.getFullYear();
  var m = current.getMonth() - birthDate.getMonth();

  if (m < 0 || (m === 0 && current.getDate() < birthDate.getDate())) {
    age--;
    m += 12; // 誕生日を迎えていない場合は、月を調整
  }

  // 月を基にした年齢の小数部分を計算
  var ageWithMonths = age + m / 12;

  // 小数点第一位までの年齢を返す
  return Math.round(ageWithMonths * 10) / 10;
}

function make_sheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //初期設定
  headers_for_input = [["ローン金額",	"年利率",	"月々の返済額",	"開始年",	"開始月",	"生年月日"]];
  sheet.getRange("A1:F1").setValues(headers_for_input);
}

function calculateLoanRepaymentWithAge(loanAmount, annualInterestRate, monthlyRepayment, startYear, startMonth, birthdate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 
  // A4行目以降をクリアする
  var lastRow = sheet.getLastRow();
  if (lastRow >= 4) {
    // A4行目からシートの最終行までをクリア
    sheet.getRange(4, 1, lastRow - 3, 7).clearContent();
  }

  var monthlyInterestRate = annualInterestRate / 12 / 100; // 月利率を計算
  var data = []; // 返済データを格納する配列

  var date = new Date(startYear, startMonth - 1, 1); // スタート年月をDateオブジェクトとして設定
  var month = 1; // ここでmonth変数を定義しています

  // 計算結果のヘッダー
  var headers = [["年月", "月", "返済額", "利息", "元本", "残高", "年齢"]];

  while (loanAmount > 0) {
    var interest = loanAmount * monthlyInterestRate; // その月の利息を計算
    var principal = monthlyRepayment - interest; // その月の元本を計算
    if(principal > loanAmount) {
      principal = loanAmount;
      monthlyRepayment = interest + principal;
    }
    loanAmount -= principal; // 残高を更新

    var yearMonth = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM");
    var age = calculateAge(birthdate, date); // その返済月の年齢を計算
    data.push([yearMonth, month++, monthlyRepayment, interest, principal, loanAmount, age]); // データ配列に追加

    date.setMonth(date.getMonth() + 1); // 次の月へ更新

    if(principal <= 0) {
      break; // 残高が0以下になったらループを終了
    }
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // 計算結果のヘッダーをA3行目に設定
  sheet.getRange("A3:G3").setValues(headers);
  // データの書き込みをA4セルから開始する
  if (data.length > 0) {
    sheet.getRange(4, 1, data.length, data[0].length).setValues(data); // A4セルからデータを一度に追加
  }
}

function testCalculateLoanRepaymentWithStartDate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // スプレッドシートから値を読み取る
  var loanAmount = sheet.getRange("A2").getValue(); // ローン金額
  var annualInterestRate = sheet.getRange("B2").getValue(); // 年利率
  var monthlyRepayment = sheet.getRange("C2").getValue(); // 月々の返済額
  var startYear = sheet.getRange("D2").getValue(); // 開始年
  var startMonth = sheet.getRange("E2").getValue(); // 開始月
  var birthdate = sheet.getRange("F2").getValue(); // 生年月日
  
  calculateLoanRepaymentWithAge(loanAmount, annualInterestRate, monthlyRepayment, startYear, startMonth, birthdate);
}
