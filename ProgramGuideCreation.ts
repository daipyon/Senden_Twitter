import RPA from 'ts-rpa';
const fs = require('fs');
const request = require('request');
var formatCSV = '';

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Twitter 番組表作成】サムネのダウンロードが完了しました`];

// スプレッドシートIDとシート名を記載
// const mySSID = process.env.My_SheetID;
const SSID = process.env.Senden_Twitter_SheetID;
const SSName = process.env.Senden_Twitter_SheetName;
// 作業するスプレッドシートから読み込む行数を記載
const StartRow = 6;
const LastRow = 30000;

// 番組表リンク
const Link1 = '番組表リンク①';
const Link2 = '番組表リンク②';
const Link3 = '番組表リンク③';

// 画像などを保存するフォルダのパスを記載
const DLFolder = __dirname + '/DL/';

// 作業対象行とデータを取得
const WorkData = [];

// エラー発生時のテキストを格納
const ErrorText = [];

async function Start() {
  if (ErrorText.length == 0) {
    // デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
    RPA.Logger.level = 'INFO';
    await RPA.Google.authorize({
      //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    // 番組表下書きシートの投稿日を取得
    const PostedDate = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!T3:T3`
    });
    // C列を取得
    const JudgeData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!C${StartRow}:C${LastRow}`
    });
    for (let i in JudgeData) {
      if (PostedDate[0][0] == JudgeData[i][0]) {
        await RPA.Logger.info(
          '投稿日の日付　　　　　   　→ ',
          PostedDate[0][0]
        );
        await RPA.Logger.info('C列の日付 　　　　　　　 　→ ', JudgeData[i][0]);
        const Row = Number(i) + 6;
        await RPA.Logger.info('この行のデータを取得します → ', Row);
        await Work(Row);
        break;
      }
    }
  }
  // エラー発生時の処理
  if (ErrorText.length >= 1) {
    // const DOM = await RPA.WebBrowser.driver.getPageSource();
    // await RPA.Logger.info(DOM);
    await RPA.SystemLogger.error(ErrorText);
    Slack_Text[0] = `【Twitter 番組表作成】でエラー発生しました\n${ErrorText}`;
    await RPA.WebBrowser.takeScreenshot();
  }
  await RPA.Logger.info('作業を終了します');
  await SlackPost(Slack_Text[0]);
  await RPA.WebBrowser.quit();
  await RPA.sleep(1000);
  await process.exit();
}

Start();

async function Work(Row) {
  try {
    await RPA.Google.authorize({
      // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    // 日付の一致判定
    await JudgeData(WorkData, Row, Link1, Link2, Link3);
    // フォトショ用
    await PhotoshopData(WorkData, Row);
  } catch (error) {
    ErrorText[0] = error;
    await Start();
  }
}

async function SlackPost(Text) {
  await RPA.Slack.chat.postMessage({
    token: Slack_Token,
    channel: Slack_Channel,
    text: `${Text}`
  });
}

async function JudgeData(WorkData, Row, Link1, Link2, Link3) {
  // 概要ページに遷移
  for (let i = 7; i <= 9; i++) {
    // 番組表下書きシートのデータ(B〜AP列)を取得
    WorkData[0] = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!B${Row}:AP${Row}`
    });
    await RPA.Logger.info(WorkData[0]);
    if (i == 7) {
      await RPA.Logger.info(`${Link1} の日付一致判定を開始します`);
    }
    if (i == 8) {
      await RPA.Logger.info(`${Link2} の日付一致判定を開始します`);
    }
    if (i == 9) {
      await RPA.Logger.info(`${Link3} の日付一致判定を開始します`);
    }
    await RPA.WebBrowser.get(WorkData[0][0][i]);
    await RPA.sleep(2000);
    // 番組の日付・曜日・開始時間・終了時間を取得
    const PageDate = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className: 'com-tv-SlotHeader__air-time'
      }),
      5000
    );
    const PageDateText = await PageDate.getText();
    const PageDateText2 = await PageDateText.split(/[()〜]/);
    const OnAirDate = await PageDateText2[0].replace('月', '/').slice(0, -1);
    // 日付の一致判定
    if (OnAirDate == WorkData[0][0][0]) {
      await RPA.Logger.info('日付が一致です');
      const StartTime = await PageDateText2[2].replace(/\s+/g, '');
      await RPA.Logger.info('開始時間  　　　　　　　 　→ ', StartTime);
      const PageDateText3 = await StartTime.split(/[:]/);
      const PageDateText4 = await OnAirDate.split(/[/]/);
      var m = await ('0' + PageDateText4[0]).slice(-2);
      var d = await ('0' + PageDateText4[1]).slice(-2);
      var d1 = Number(d) + 1;
      // 翌日
      var d2 = await ('0' + d1).slice(-2);
      var y = await WorkData[0][0][1].slice(0, -5);
      var y2 = await y.slice(-2);
      var dateAttr = new Date(`${y}-${m}-${d}`);
      var monthAtrr = await dateAttr.getMonth();
      var mNames = [
        'Jan',
        'Feb',
        'Mar',
        'Apr',
        'May',
        'Jun',
        'Jul',
        'Aug',
        'Sep',
        'Oct',
        'Nov',
        'Dec'
      ];
      var month = mNames[monthAtrr];
      var mS = await month.toString();
      // 時間帯
      const date = new Date(`${d}-${mS}-${y2} ${StartTime}:00 GMT`);
      // 朝
      const T4_00 = new Date(`${d}-${mS}-${y2} 04:00:00 GMT`);
      const T8_59 = new Date(`${d}-${mS}-${y2} 08:59:00 GMT`);
      // 午前
      const T9_00 = new Date(`${d}-${mS}-${y2} 09:00:00 GMT`);
      const T11_59 = new Date(`${d}-${mS}-${y2} 11:59:00 GMT`);
      // 昼
      const T12_00 = new Date(`${d}-${mS}-${y2} 12:00:00 GMT`);
      const T12_59 = new Date(`${d}-${mS}-${y2} 12:59:00 GMT`);
      // 午後
      const T13_00 = new Date(`${d}-${mS}-${y2} 13:00:00 GMT`);
      const T17_59 = new Date(`${d}-${mS}-${y2} 17:59:00 GMT`);
      // 夜
      const T18_00 = new Date(`${d}-${mS}-${y2} 18:00:00 GMT`);
      const T23_59 = new Date(`${d}-${mS}-${y2} 23:59:00 GMT`);
      // 今夜
      const T24_00 = new Date(`${d}-${mS}-${y2} 24:00:00 GMT`);
      const T26_59 = new Date(`${d2}-${mS}-${y2} 02:59:00 GMT`);
      const Asa = date >= T4_00 && date <= T8_59;
      const Gozen = date >= T9_00 && date <= T11_59;
      const Hiru = date >= T12_00 && date <= T12_59;
      const Gogo = date >= T13_00 && date <= T17_59;
      const Yoru = date >= T18_00 && date <= T23_59;
      const Konya = date >= T24_00 && date <= T26_59;
      // コピーライトを取得
      const CopyRight = await RPA.WebBrowser.findElementByClassName(
        'com-o-Footer__copyright'
      );

      // テスト用
      // const CopyRightText = '(C)2018MBC';

      // 本番用
      const CopyRightText = await CopyRight.getText();

      await RPA.Logger.info(CopyRightText);
      // 曜日の一致判定
      const date1 = new Date(
        `${d}-${mS}-${y2} ${PageDateText3[0]}:${PageDateText3[1]}:00 GMT+0900`
      );
      var w = await date1.getDay();
      var w2 = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'][w];
      await RPA.Logger.info('シートの曜日  　　　　　 　→ ', WorkData[0][0][2]);
      await RPA.Logger.info('変換した曜日  　　　　　 　→ ', w2);
      if (w2 == WorkData[0][0][2]) {
        await RPA.Logger.info('曜日も一致です');
        if (i == 7) {
          // 番組表リンク①の入力処理
          await Link1SetValue(
            Row,
            w2,
            Asa,
            Gozen,
            Hiru,
            Gogo,
            Yoru,
            Konya,
            PageDateText3,
            StartTime,
            CopyRightText
          );
        }
        if (i == 8) {
          // 番組表リンク②の入力処理
          await Link2SetValue(
            Row,
            Asa,
            Gozen,
            Hiru,
            Gogo,
            Yoru,
            Konya,
            PageDateText3,
            StartTime,
            CopyRightText
          );
        }
        if (i == 9) {
          // 番組表リンク③の入力処理
          await Link3SetValue(
            Row,
            Asa,
            Gozen,
            Hiru,
            Gogo,
            Yoru,
            Konya,
            PageDateText3,
            StartTime,
            CopyRightText
          );
        }
        // 名前を変更したサムネを保持する変数
        const Renamming = [];
        // サムネをダウンロード
        await ThumbnailDownload(Renamming);
        if (i == 7) {
          // AL列に画像名を記載
          await RenammingSetValue(`AL${Row}`, Renamming);
        }
        if (i == 8) {
          // AM列に画像名を記載
          await RenammingSetValue(`AM${Row}`, Renamming);
        }
        if (i == 9) {
          // AN列に画像名を記載
          await RenammingSetValue(`AN${Row}`, Renamming);
        }
        // 指定したGoogleDriveのIDにアップロード
        await GoogleUpload(Renamming);
      } else {
        await RPA.Logger.info('曜日が不一致のため記載ミスです');
        if (i == 7) {
          await RPA.Logger.info(`${Link1}の "曜日" が不一致です`);
          const ErrorText = `【${Link1}】\n\n "曜日" が不一致です`;
          // AB列にエラー文言を記載
          await ErrorSetValue(`AB${Row}`, ErrorText);
        }
        if (i == 8) {
          await RPA.Logger.info(`${Link2}の "曜日" が不一致です`);
          const ErrorText = `【${Link2}】\n\n "曜日" が不一致です`;
          // 先にエラー文言が入っていた場合、書き換え後の文言を記載
          if (WorkData[0][0][26] != undefined) {
            await ErrorSetValue2(`AB${Row}`, ErrorText);
          } else {
            // AB列にエラー文言を記載
            await ErrorSetValue(`AB${Row}`, ErrorText);
          }
        }
        if (i == 9) {
          await RPA.Logger.info(`${Link3}の "曜日" が不一致です`);
          const ErrorText = `【${Link3}】\n\n "曜日" が不一致です`;
          // 先にエラー文言が入っていた場合、書き換え後の文言を記載
          if (WorkData[0][0][26] != undefined) {
            await ErrorSetValue2(`AB${Row}`, ErrorText);
          } else {
            // AB列にエラー文言を記載
            await ErrorSetValue(`AB${Row}`, ErrorText);
          }
        }
      }
    } else {
      await RPA.Logger.info('日付が不一致のため、開始時間で判定します');
      const StartTime = await PageDateText2[2].replace(/\s+/g, '');
      await RPA.Logger.info('開始時間  　　　　　　　 　→ ', StartTime);
      const PageDateText3 = await StartTime.split(/[:]/);
      const PageDateText4 = await OnAirDate.split(/[/]/);
      var m = await ('0' + PageDateText4[0]).slice(-2);
      var d = await ('0' + PageDateText4[1]).slice(-2);
      var d1 = Number(d) + 1;
      // 翌日
      var d2 = await ('0' + d1).slice(-2);
      var y = await WorkData[0][0][1].slice(0, -5);
      var y2 = await y.slice(-2);
      var dateAttr = new Date(`${y}-${m}-${d}`);
      var monthAtrr = await dateAttr.getMonth();
      var mNames = [
        'Jan',
        'Feb',
        'Mar',
        'Apr',
        'May',
        'Jun',
        'Jul',
        'Aug',
        'Sep',
        'Oct',
        'Nov',
        'Dec'
      ];
      var month = mNames[monthAtrr];
      var mS = await month.toString();
      // 時間帯
      const date = new Date(`${d}-${mS}-${y2} ${StartTime}:00 GMT`);
      // 朝
      const T4_00 = new Date(`${d}-${mS}-${y2} 04:00:00 GMT`);
      const T8_59 = new Date(`${d}-${mS}-${y2} 08:59:00 GMT`);
      // 午前
      const T9_00 = new Date(`${d}-${mS}-${y2} 09:00:00 GMT`);
      const T11_59 = new Date(`${d}-${mS}-${y2} 11:59:00 GMT`);
      // 昼
      const T12_00 = new Date(`${d}-${mS}-${y2} 12:00:00 GMT`);
      const T12_59 = new Date(`${d}-${mS}-${y2} 12:59:00 GMT`);
      // 午後
      const T13_00 = new Date(`${d}-${mS}-${y2} 13:00:00 GMT`);
      const T17_59 = new Date(`${d}-${mS}-${y2} 17:59:00 GMT`);
      // 夜
      const T18_00 = new Date(`${d}-${mS}-${y2} 18:00:00 GMT`);
      const T23_59 = new Date(`${d}-${mS}-${y2} 23:59:00 GMT`);
      // 今夜
      const T24_00 = new Date(`${d}-${mS}-${y2} 24:00:00 GMT`);
      const T26_59 = new Date(`${d2}-${mS}-${y2} 02:59:00 GMT`);
      const Asa = date >= T4_00 && date <= T8_59;
      const Gozen = date >= T9_00 && date <= T11_59;
      const Hiru = date >= T12_00 && date <= T12_59;
      const Gogo = date >= T13_00 && date <= T17_59;
      const Yoru = date >= T18_00 && date <= T23_59;
      const Konya = date <= T24_00 && date <= T26_59;
      // コピーライトを取得
      const CopyRight = await RPA.WebBrowser.findElementByClassName(
        'com-o-Footer__copyright'
      );

      // テスト用
      // const CopyRightText = '(C)2018MBC';

      // 本番用
      const CopyRightText = await CopyRight.getText();
      await RPA.Logger.info(CopyRightText);
      // 開始時間の比較
      const date1 = new Date(
        `${d}-${mS}-${y2} ${PageDateText3[0]}:${PageDateText3[1]}:00 GMT`
      );
      const date2 = new Date(`${d}-${mS}-${y2} 02:00:00 GMT`);
      await RPA.Logger.info('シートの開始時間  　　　 　→ ', date1);
      await RPA.Logger.info('変換した開始時間  　　　 　→ ', date2);
      if (date1 < date2) {
        await RPA.Logger.info('開始時間が2:00までなので合っています');
        if (i == 7) {
          // 番組表リンク①の入力処理
          await Link1SetValue(
            Row,
            w2,
            Asa,
            Gozen,
            Hiru,
            Gogo,
            Yoru,
            Konya,
            PageDateText3,
            StartTime,
            CopyRightText
          );
        }
        if (i == 8) {
          // 番組表リンク②の入力処理
          await Link2SetValue(
            Row,
            Asa,
            Gozen,
            Hiru,
            Gogo,
            Yoru,
            Konya,
            PageDateText3,
            StartTime,
            CopyRightText
          );
        }
        if (i == 9) {
          // 番組表リンク③の入力処理
          await Link3SetValue(
            Row,
            Asa,
            Gozen,
            Hiru,
            Gogo,
            Yoru,
            Konya,
            PageDateText3,
            StartTime,
            CopyRightText
          );
        }
        // 名前を変更したサムネを保持する変数
        const Renamming = [];
        // サムネをダウンロード
        await ThumbnailDownload(Renamming);
        if (i == 7) {
          // AL列に画像名を記載
          await RenammingSetValue(`AL${Row}`, Renamming);
        }
        if (i == 8) {
          // AM列に画像名を記載
          await RenammingSetValue(`AM${Row}`, Renamming);
        }
        if (i == 9) {
          // AN列に画像名を記載
          await RenammingSetValue(`AN${Row}`, Renamming);
        }
        // 指定したGoogleDriveのIDにアップロード
        await GoogleUpload(Renamming);
      } else {
        await RPA.Logger.info('日付が不一致のため記載ミスです');
        if (i == 7) {
          await RPA.Logger.info(`${Link1}の "日付" が不一致です`);
          const ErrorText = `【${Link1}】\n\n "日付" が不一致です`;
          // AB列にエラー文言を記載
          await ErrorSetValue(`AB${Row}`, ErrorText);
        }
        if (i == 8) {
          await RPA.Logger.info(`${Link2}の "日付" が不一致です`);
          const ErrorText = `【${Link2}】\n\n "日付" が不一致です`;
          // 先にエラー文言が入っていた場合、書き換え後の文言を記載
          if (WorkData[0][0][26] != undefined) {
            await ErrorSetValue2(`AB${Row}`, ErrorText);
          } else {
            // AB列にエラー文言を記載
            await ErrorSetValue(`AB${Row}`, ErrorText);
          }
        }
        if (i == 9) {
          await RPA.Logger.info(`${Link3}の "日付" が不一致です`);
          const ErrorText = `【${Link3}】\n\n "日付" が不一致です`;
          // 先にエラー文言が入っていた場合、書き換え後の文言を記載
          if (WorkData[0][0][26] != undefined) {
            await ErrorSetValue2(`AB${Row}`, ErrorText);
          } else {
            // AB列にエラー文言を記載
            await ErrorSetValue(`AB${Row}`, ErrorText);
          }
        }
      }
    }
  }
}

async function Link1SetValue(
  Row,
  w2,
  Asa,
  Gozen,
  Hiru,
  Gogo,
  Yoru,
  Konya,
  PageDateText3,
  StartTime,
  CopyRightText
) {
  const DateSplit = await WorkData[0][0][0].split(/[/]/);
  // AC列に月を入力
  await MonthSetValue(`AC${Row}`, DateSplit);
  // AD列に日を入力
  await DaySetValue(`AD${Row}`, DateSplit);
  // AE列に曜日を入力
  await WeekdaySetValue(`AE${Row}`, w2);
  // AF列に時間帯を入力
  if (Asa) {
    await AsaSetValue(`AF${Row}`);
  }
  if (Gozen) {
    await GozenSetValue(`AF${Row}`);
  }
  if (Hiru) {
    await HiruSetValue(`AF${Row}`);
  }
  if (Gogo) {
    await GogoSetValue(`AF${Row}`);
  }
  if (Yoru) {
    await YoruSetValue(`AF${Row}`);
  }
  if (Konya) {
    await KonyaSetValue(`AF${Row}`);
  }
  // AG列に開始時間を入力
  await StartTimeSetValue(PageDateText3, StartTime, `AG${Row}`);
  // AO列にコピーライトを入力
  await CopyRightSetValue(CopyRightText, `AO${Row}`);
}

async function Link2SetValue(
  Row,
  Asa,
  Gozen,
  Hiru,
  Gogo,
  Yoru,
  Konya,
  PageDateText3,
  StartTime,
  CopyRightText
) {
  // AH列に時間帯を入力
  if (Asa) {
    await AsaSetValue(`AH${Row}`);
  }
  if (Gozen) {
    await GozenSetValue(`AH${Row}`);
  }
  if (Hiru) {
    await HiruSetValue(`AH${Row}`);
  }
  if (Gogo) {
    await GogoSetValue(`AH${Row}`);
  }
  if (Yoru) {
    await YoruSetValue(`AH${Row}`);
  }
  if (Konya) {
    await KonyaSetValue(`AH${Row}`);
  }
  // AI列に開始時間を入力
  await StartTimeSetValue(PageDateText3, StartTime, `AI${Row}`);
  // AO列にコピーライトを入力
  await CopyRightSetValue(CopyRightText, `AO${Row}`);
}

async function Link3SetValue(
  Row,
  Asa,
  Gozen,
  Hiru,
  Gogo,
  Yoru,
  Konya,
  PageDateText3,
  StartTime,
  CopyRightText
) {
  // AJ列に時間帯を入力
  if (Asa) {
    await AsaSetValue(`AJ${Row}`);
  }
  if (Gozen) {
    await GozenSetValue(`AJ${Row}`);
  }
  if (Hiru) {
    await HiruSetValue(`AJ${Row}`);
  }
  if (Gogo) {
    await GogoSetValue(`AJ${Row}`);
  }
  if (Yoru) {
    await YoruSetValue(`AJ${Row}`);
  }
  if (Konya) {
    await KonyaSetValue(`AJ${Row}`);
  }
  // AK列に開始時間を入力
  await StartTimeSetValue(PageDateText3, StartTime, `AK${Row}`);
  // AO列にコピーライトを入力
  await CopyRightSetValue(CopyRightText, `AO${Row}`);
}

async function PhotoshopData(WorkData, Row) {
  // AP列にPSD保存用の画像名を記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!AP${Row}:AP${Row}`,
    values: [[`${WorkData[0][0][1]}.psd`]]
  });
  // Photoshopで使用するデータ(AC〜AP列)を取得
  const PhotoshopData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!AC${Row}:AP${Row}`
  });
  await RPA.Logger.info(PhotoshopData);
  // 取得したデータをcsvで保存（この関数内で書かないとファイルに記載されない）
  await exportCSV(PhotoshopData);
  async function exportCSV(content) {
    for (var i = 0; i < content.length; i++) {
      var value = content[i];
      for (var j = 0; j < value.length; j++) {
        var innerValue = value[j] === null ? '' : value[j].toString();
        var result = await innerValue.replace(/"/g, '""');
        if (result.search(/("|,|\n)/g) >= 0) result = '"' + result + '"';
        if (j > 0) formatCSV += ',';
        formatCSV += result;
      }
      formatCSV += '\n';
    }
    await fs.writeFile('formList.csv', formatCSV, 'utf8', function(err) {
      if (err) {
        RPA.Logger.info('保存できませんでした');
      } else {
        RPA.Logger.info('保存しました');
      }
    });
  }
}

// 月を記載する関数
async function MonthSetValue(Row, DateSplit) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [[DateSplit[0]]]
  });
}

// 日を記載する関数
async function DaySetValue(Row, DateSplit) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [[DateSplit[1]]]
  });
}

// 曜日を記載する関数
async function WeekdaySetValue(Row, w2) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [[w2]]
  });
}

// "朝"と記載する関数
async function AsaSetValue(Row) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [['朝']]
  });
}

// "午前"と記載する関数
async function GozenSetValue(Row) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [['午前']]
  });
}

// "昼"と記載する関数
async function HiruSetValue(Row) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [['昼']]
  });
}

// "午後"と記載する関数
async function GogoSetValue(Row) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [['午後']]
  });
}

// "夜"と記載する関数
async function YoruSetValue(Row) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [['夜']]
  });
}

// "今夜"と記載する関数
async function KonyaSetValue(Row) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [['今夜']]
  });
}

// 開始時間を記載する関数
async function StartTimeSetValue(PageDateText3, StartTime, Row) {
  if (Number(PageDateText3[0]) > 12) {
    const ampm = Number(PageDateText3[0]) - 12;
    const ampmStartTime = ampm + ':' + PageDateText3[1];
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!${Row}:${Row}`,
      values: [[`${ampmStartTime}〜`]]
    });
  } else {
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!${Row}:${Row}`,
      values: [[`${StartTime}〜`]]
    });
  }
}

// コピーライトを記載する関数
async function CopyRightSetValue(CopyRightText, Row) {
  // AO列にコピーライトを入力
  if (
    CopyRightText.indexOf('Abema') > -1 ||
    CopyRightText.indexOf('abema') > -1 ||
    CopyRightText.indexOf('ABEMA') > -1
  ) {
    // 記載をスルー
    await RPA.Logger.info('AbemaTVです');
  } else {
    await RPA.Logger.info('AbemaTV以外です');
    // 「(C)」もしくは「制作著作」の文字が含まれている場合、「©」に変換して記載
    if (CopyRightText.indexOf('(C)') > -1) {
      const CopyRightTextReplace = await CopyRightText.replace('(C)', '©');
      // 先にコピーライトが入っていた場合、書き換え後の文言を記載
      if (WorkData[0][0][39] != undefined) {
        await RPA.Logger.info('コピーライトが入っているため追記します');
        const NewCopyRightText =
          WorkData[0][0][39] + ` ${CopyRightTextReplace}`;
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!${Row}:${Row}`,
          values: [[NewCopyRightText]]
        });
      } else {
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!${Row}:${Row}`,
          values: [[CopyRightTextReplace]]
        });
      }
    } else if (CopyRightText.indexOf('制作著作') > -1) {
      const CopyRightTextReplace = await CopyRightText.replace('制作著作', '©');
      // 先にコピーライトが入っていた場合、書き換え後の文言を記載
      if (WorkData[0][0][39] != undefined) {
        await RPA.Logger.info('コピーライトが入っているため追記します');
        const NewCopyRightText =
          WorkData[0][0][39] + ` ${CopyRightTextReplace}`;
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!${Row}:${Row}`,
          values: [[NewCopyRightText]]
        });
      } else {
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!${Row}:${Row}`,
          values: [[CopyRightTextReplace]]
        });
      }
    } else {
      // 先にコピーライトが入っていた場合、書き換え後の文言を記載
      if (WorkData[0][0][39] != undefined) {
        await RPA.Logger.info('コピーライトが入っているため追記します');
        const NewCopyRightText = WorkData[0][0][36] + ` ${Text}`;
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!${Row}:${Row}`,
          values: [[NewCopyRightText]]
        });
      } else {
        await RPA.Google.Spreadsheet.setValues({
          spreadsheetId: `${SSID}`,
          range: `${SSName}!${Row}:${Row}`,
          values: [[CopyRightText]]
        });
      }
    }
  }
}

// サムネをダウンロードする関数
async function ThumbnailDownload(Renamming) {
  await RPA.Logger.info('サムネをダウンロードします');
  // 画像のURLを取得
  const ImageUrl = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      tagName: 'source'
    }),
    5000
  );
  const ImageUrlText = await ImageUrl.getAttribute('srcset');
  const ImageUrlSplit = await ImageUrlText.split('?');
  const ImageUrlRename = await ImageUrlSplit[0].replace(/v\d+.webp/, 'jpg');
  await RPA.Logger.info(ImageUrlRename);
  const ImageUrlRenameSplit = await ImageUrlRename.split('/');
  // 名前を変更
  Renamming[0] = await ImageUrlRenameSplit[6].replace(
    /thumb001.jpg/,
    `${ImageUrlRenameSplit[5]}.jpg`
  );
  await RPA.Logger.info(Renamming);
  await request(
    { method: 'GET', url: ImageUrlRename, encoding: null },
    async function(error, response, body) {
      if (!error && response.statusCode === 200) {
        await fs.writeFileSync(`${DLFolder}/${Renamming}`, body, 'binary');
      }
    }
  );
}

// Google Driveへアップロードする関数
async function GoogleUpload(Renamming) {
  await RPA.Logger.info('アップロード実行中...');
  await RPA.sleep(3000);
  await RPA.Google.Drive.upload({
    filename: Renamming[0],
    parents: [`${process.env.Senden_Twitter_GoogleDriveFolderID2}`]
  });
  await RPA.Logger.info('アップロード完了しました');
}

// 画像名を記載する関数
async function RenammingSetValue(Row, Renamming) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [[Renamming[0]]]
  });
}

// エラー文言を記載する関数
async function ErrorSetValue(Row, ErrorText) {
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [[ErrorText]]
  });
}

async function ErrorSetValue2(Row, ErrorText) {
  await RPA.Logger.info('エラー文言が入っているため追記します');
  const NewErrorText =
    WorkData[0][0][26] + `\n\n----------------------------\n\n${ErrorText}`;
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!${Row}:${Row}`,
    values: [[NewErrorText]]
  });
}
