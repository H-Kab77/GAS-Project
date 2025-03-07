function saveSpreadsheetDataWithTimestamp() {
    var ss = SpreadsheetApp.openById('1z3F0b9aaRSowqnXQQzrK0y8mD4DMZqbAIVqW6cbS_x4');
    var sheet = ss.getSheetByName('データ一覧');
    var data = sheet.getDataRange().getValues();

    // 日時を取得（YYYY/MM/DD HH:mm形式）
    var now = new Date();
    var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

    // 先頭に「取得日時」を追加した新しいデータを作る
    var newData = data.map(function(row, index) {
        if (index === 0) {
            return ['取得日時'].concat(row);  // ヘッダー行に「取得日時」を追加
        } else {
            return [timestamp].concat(row);   // 各データ行に取得日時を追加
        }
    });

    // 保存先シート（保存データ）に保存
    var saveSheet = ss.getSheetByName('保存データ');
    if (!saveSheet) {
        saveSheet = ss.insertSheet('保存データ');
    }
    saveSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);

    Logger.log('データ保存（日時付き）完了！');

    // メール送信
    var emailAddress = "your-email@example.com";  // メールを送りたいアドレス
    var subject = "GASデータ取得完了";  // 件名
    var body = "データ取得が完了しました。以下のリンクからデータを確認できます。";  // 本文
    MailApp.sendEmail(emailAddress, subject, body);  // メール送信

    // Googleドライブの指定フォルダにCSVファイル保存
    var folder = DriveApp.getFolderById('1VuiMO55IWtzBCFUT77Xtrp2yWqEFYd9u');  // 先ほどのフォルダID
    var fileName = 'データ_' + timestamp + '.csv';
    var csvFile = Utilities.formatString('%s', newData.map(function(row) {
        return row.join(',');
    }).join('\n'));
    folder.createFile(fileName, csvFile, MimeType.CSV);  // 指定フォルダにCSV保存
}
