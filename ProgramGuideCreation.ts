import RPA from 'ts-rpa';
import { WebDriver, By, FileDetector, Key } from 'selenium-webdriver';
import { rootCertificates } from 'tls';
import { worker } from 'cluster';
import { cachedDataVersionTag } from 'v8';
import { start } from 'repl';
import { Command } from 'selenium-webdriver/lib/command';
import { drive } from "googleapis/build/src/apis/drive";
import { Driver } from 'selenium-webdriver/safari';
import { Dataset } from '@google-cloud/bigquery';
var fs = require('fs');
var request = require('request');
var formatCSV = '';
// デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
RPA.Logger.level = 'INFO';

// スプレッドシートIDとシート名を記載
const SSID = process.env.Senden_Twitter_SheetID;
const SSName = process.env.Senden_Twitter_SheetName;
// 画像などを保存するフォルダのパスを記載 ※.envファイルは同じにしない
const DownloadFolder = process.env.Senden_Twitter_DownloadFolder;
// 番組表リンク
const Link1 = process.env.Program_Guide_Link1;
const Link2 = process.env.Program_Guide_Link2;
const Link3 = process.env.Program_Guide_Link3;
// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;

async function WorkStart() {
  await RPA.Google.authorize({
    // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });

  // 作業対象行とデータを取得
  const Row = [];
  await GetDataRow(Row);

  // 日付の一致判定
  const WorkData = [];
//   await JudgeData(WorkData, Row, Link1, Link2, Link3);

//   await PhotoshopData(WorkData, Row);
}

async function Start() {
  try {
    await RPA.Google.authorize({
      //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    await WorkStart();
  } catch (error) {
    RPA.SystemLogger.error(error);
    await RPA.WebBrowser.takeScreenshot();
    // Slackにも通知
    await RPA.Slack.chat.postMessage({
      token: Slack_Token,
      channel: Slack_Channel,
      text: '【宣伝_Twitter 番組表作成】（100-00422用）でエラーが発生しました！'
    });
  }
  await RPA.WebBrowser.quit();
}

Start();

async function GetDataRow(Row) {
  // 番組表下書きシートの投稿日を取得
  const PostedDate = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!T3:T3`
  });
  // C列を取得
  const JudgeData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!C6:C30000`
  });
  for (let i in JudgeData) {
    if (PostedDate[0][0] == JudgeData[i][0]) {
      RPA.Logger.info(JudgeData[i][0]);
      Row[0] = Number(i) + 6;
      break;
    }
  }
  RPA.Logger.info('この行のデータを取得します → ', Row[0]);
}

async function JudgeData(WorkData, Row, Link1, Link2, Link3) {
  // 概要ページに遷移
  for (var i = 7; i <= 9; i++) {
    // 番組表下書きシートのデータ(B〜AP列)を取得
    WorkData[0] = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName}!B${Row[0]}:AP${Row[0]}`
    });
    RPA.Logger.info(WorkData[0]);
    if (i == 7) {
      RPA.Logger.info(`${Link1}の日付一致判定を開始します`);
    }
    if (i == 8) {
      RPA.Logger.info(`${Link2}の日付一致判定を開始します`);
    }
    if (i == 9) {
      RPA.Logger.info(`${Link3}の日付一致判定を開始します`);
    }
    await RPA.WebBrowser.get(WorkData[0][0][i]);
    await RPA.sleep(2000);
    // 番組の日付・曜日・開始時間・終了時間を取得
    const PageDate = await RPA.WebBrowser.findElementByClassName(
      'com-tv-SlotHeader__air-time'
    );
    const PageDateText = await PageDate.getText();
    const PageDateText2 = PageDateText.split(/[()〜]/);
    const OnAirDate = PageDateText2[0].replace('月', '/').slice(0, -1);
    // 日付の一致判定
    if (OnAirDate == WorkData[0][0][0]) {
      RPA.Logger.info('日付が一致です');
      const StartTime = PageDateText2[2].replace(/\s+/g, '');
      RPA.Logger.info(StartTime);
      const PageDateText3 = StartTime.split(/[:]/);
      const PageDateText4 = OnAirDate.split(/[/]/);
      var m = ('0' + PageDateText4[0]).slice(-2);
      var d = ('0' + PageDateText4[1]).slice(-2);
      var d1 = Number(d) + 1;
      // 翌日
      var d2 = ('0' + d1).slice(-2);
      var y = WorkData[0][0][1].slice(0, -5);
      var y2 = y.slice(-2);
      var dateAttr = new Date(`${y}-${m}-${d}`);
      var monthAtrr = dateAttr.getMonth();
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
      var mS = month.toString();
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
      // コピーライトを取得
      const CopyRight = await RPA.WebBrowser.findElementByClassName(
        'com-o-Footer__copyright'
      );
      const CopyRightText = await CopyRight.getText();
      RPA.Logger.info(CopyRightText);
      // 曜日の一致判定
      const date1 = new Date(
        `${d}-${mS}-${y2} ${PageDateText3[0]}:${PageDateText3[1]}:00 GMT+0900`
      );
      var w = date1.getDay();
      var w2 = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'][w];
      RPA.Logger.info(w2);
      RPA.Logger.info(WorkData[0][0][2]);
      if (w2 == WorkData[0][0][2]) {
        RPA.Logger.info('曜日も一致です');
        if (i == 7) {
          const DateSplit = WorkData[0][0][0].split(/[/]/);
          // AC列に月を入力
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AC${Row[0]}:AC${Row[0]}`,
            values: [[DateSplit[0]]]
          });
          // AD列に日を入力
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AD${Row[0]}:AD${Row[0]}`,
            values: [[DateSplit[1]]]
          });
          // AE列に曜日を入力
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AE${Row[0]}:AE${Row[0]}`,
            values: [[w2]]
          });
          // AF列に時間帯を入力
          if (date >= T4_00 && date <= T8_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['朝']]
            });
          }
          if (date >= T9_00 && date <= T11_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['午前']]
            });
          }
          if (date >= T12_00 && date <= T12_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['昼']]
            });
          }
          if (date >= T13_00 && date <= T17_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['午後']]
            });
          }
          if (date >= T18_00 && date <= T23_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['夜']]
            });
          }
          if (date >= T24_00 && date <= T26_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['今夜']]
            });
          }
          // AG列に開始時間を入力
          if (Number(PageDateText3[0]) > 12) {
            const ampm = Number(PageDateText3[0]) - 12;
            const ampmStartTime = ampm + ':' + PageDateText3[1];
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AG${Row[0]}:AG${Row[0]}`,
              values: [[`${ampmStartTime}〜`]]
            });
          } else {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AG${Row[0]}:AG${Row[0]}`,
              values: [[`${Number(PageDateText3[0])}:${PageDateText3[1]}〜`]]
            });
          }
          // AO列にコピーライトを入力
          if (CopyRightText.indexOf('Abema') > -1 || CopyRightText.indexOf('abema') > -1 || CopyRightText.indexOf('ABEMA') > -1) {
            RPA.Logger.info('AbemaTVです');
          } else {
            RPA.Logger.info('AbemaTV以外です');
            // 「(C)」もしくは「制作著作」の文字が含まれている場合、「©」に変換して記載
            if (CopyRightText.indexOf('(C)') > -1) {
              const CopyRightTextReplace = CopyRightText.replace('(C)', '©');
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${SSID}`,
                range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                values: [[CopyRightTextReplace]]
              });
            } else if (CopyRightText.indexOf('制作著作') > -1) {
              const CopyRightTextReplace = CopyRightText.replace(
                '制作著作',
                '©'
              );
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${SSID}`,
                range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                values: [[CopyRightTextReplace]]
              });
            } else {
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${SSID}`,
                range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                values: [[CopyRightText]]
              });
            }
          }
        }
        if (i == 8) {
          // AH列に時間帯を入力
          if (date >= T4_00 && date <= T8_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['朝']]
            });
          }
          if (date >= T9_00 && date <= T11_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['午前']]
            });
          }
          if (date >= T12_00 && date <= T12_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['昼']]
            });
          }
          if (date >= T13_00 && date <= T17_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['午後']]
            });
          }
          if (date >= T18_00 && date <= T23_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['夜']]
            });
          }
          if (date >= T24_00 && date <= T26_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['今夜']]
            });
          }
          // AI列に開始時間を入力
          if (Number(PageDateText3[0]) > 12) {
            const ampm = Number(PageDateText3[0]) - 12;
            const ampmStartTime = ampm + ':' + PageDateText3[1];
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AI${Row[0]}:AI${Row[0]}`,
              values: [[`${ampmStartTime}〜`]]
            });
          } else {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AI${Row[0]}:AI${Row[0]}`,
              values: [[`${Number(PageDateText3[0])}:${PageDateText3[1]}〜`]]
            });
          }
          // AO列にコピーライトを入力
          if (CopyRightText.indexOf('Abema') > -1 || CopyRightText.indexOf('abema') > -1 || CopyRightText.indexOf('ABEMA') > -1) {
            RPA.Logger.info('AbemaTVです');
          } else {
            RPA.Logger.info('AbemaTV以外です');
            // 「(C)」もしくは「制作著作」の文字が含まれている場合、「©」に変換して記載
            if (CopyRightText.indexOf('(C)') > -1) {
              const CopyRightTextReplace = CopyRightText.replace('(C)', '©');
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText =
                  WorkData[0][0][39] + ` ${CopyRightTextReplace}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightTextReplace]]
                });
              }
            } else if (CopyRightText.indexOf('制作著作') > -1) {
              const CopyRightTextReplace = CopyRightText.replace(
                '制作著作',
                '©'
              );
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText =
                  WorkData[0][0][39] + ` ${CopyRightTextReplace}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightTextReplace]]
                });
              }
            } else {
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText = WorkData[0][0][36] + ` ${Text}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightText]]
                });
              }
            }
          }
        }
        if (i == 9) {
          // AJ列に時間帯を入力
          if (date >= T4_00 && date <= T8_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['朝']]
            });
          }
          if (date >= T9_00 && date <= T11_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['午前']]
            });
          }
          if (date >= T12_00 && date <= T12_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['昼']]
            });
          }
          if (date >= T13_00 && date <= T17_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['午後']]
            });
          }
          if (date >= T18_00 && date <= T23_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['夜']]
            });
          }
          if (date >= T24_00 && date <= T26_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['今夜']]
            });
          }
          // AK列に開始時間を入力
          if (Number(PageDateText3[0]) > 12) {
            const ampm = Number(PageDateText3[0]) - 12;
            const ampmStartTime = ampm + ':' + PageDateText3[1];
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AK${Row[0]}:AK${Row[0]}`,
              values: [[`${ampmStartTime}〜`]]
            });
          } else {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AK${Row[0]}:AK${Row[0]}`,
              values: [[`${Number(PageDateText3[0])}:${PageDateText3[1]}〜`]]
            });
          }
          // AO列にコピーライトを入力
          if (CopyRightText.indexOf('Abema') > -1 || CopyRightText.indexOf('abema') > -1 || CopyRightText.indexOf('ABEMA') > -1) {
            RPA.Logger.info('AbemaTVです');
          } else {
            RPA.Logger.info('AbemaTV以外です');
            // 「(C)」もしくは「制作著作」の文字が含まれている場合、「©」に変換して記載
            if (CopyRightText.indexOf('(C)') > -1) {
              const CopyRightTextReplace = CopyRightText.replace('(C)', '©');
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText =
                  WorkData[0][0][39] + ` ${CopyRightTextReplace}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightTextReplace]]
                });
              }
            } else if (CopyRightText.indexOf('制作著作') > -1) {
              const CopyRightTextReplace = CopyRightText.replace(
                '制作著作',
                '©'
              );
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText =
                  WorkData[0][0][39] + ` ${CopyRightTextReplace}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightTextReplace]]
                });
              }
            } else {
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText = WorkData[0][0][36] + ` ${Text}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightText]]
                });
              }
            }
          }
        }
        RPA.Logger.info('画像をダウンロードします');
        // 画像のURLを取得
        const ImageUrlText = await RPA.WebBrowser.driver
          .findElement(
            By.css(
              '#main > div > div.c-application-DesktopAppContainer__content-container > div.c-application-DesktopAppContainer__content > main > div > div > div > div.c-tv-UpcomingSlotContainerView__header > div.c-tv-UpcomingSlotContainerView__thumbnail-column > div > div > img'
            )
          )
          .getAttribute('src');
        const ImageUrlSplit = ImageUrlText.split('?');
        const ImageUrlRename = ImageUrlSplit[0].replace(/v\d+.webp/, 'jpg');
        RPA.Logger.info(ImageUrlRename);
        const ImageUrlRenameSplit = ImageUrlRename.split('/');
        // 名前を変更
        const Renamming = ImageUrlRenameSplit[6].replace(
          /thumb001.jpg/,
          `${ImageUrlRenameSplit[5]}.jpg`
        );
        RPA.Logger.info(Renamming);
        if (i == 7) {
          // 番組表下書きシートのAL列に画像名を記載
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AL${Row[0]}:AL${Row[0]}`,
            values: [[Renamming]]
          });
        }
        if (i == 8) {
          // 番組表下書きシートのAM列に画像名を記載
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AM${Row[0]}:AM${Row[0]}`,
            values: [[Renamming]]
          });
        }
        if (i == 9) {
          // 番組表下書きシートのAN列に画像名を記載
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AN${Row[0]}:AN${Row[0]}`,
            values: [[Renamming]]
          });
        }
        // ダウンロードフォルダに保存
        request(
          { method: 'GET', url: ImageUrlRename, encoding: null },
          function(error, response, body) {
            if (!error && response.statusCode === 200) {
              fs.writeFileSync(
                `${DownloadFolder}/${Renamming}`,
                body,
                'binary'
              );
            }
          }
        );
        RPA.Logger.info('ダウンロード完了です');
      } else {
        RPA.Logger.info('曜日が不一致のため記載ミスです');
        if (i == 7) {
          RPA.Logger.info(`${Link1}の "曜日" が不一致です`);
          const ErrorText = `【${Link1}】\n\n "曜日" が不一致です`;
          // 番組表下書きシートのAB列にエラー文言を記載
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AB${Row[0]}:AB${Row[0]}`,
            values: [[ErrorText]]
          });
        }
        if (i == 8) {
          RPA.Logger.info(`${Link2}の "曜日" が不一致です`);
          const ErrorText = `【${Link2}】\n\n "曜日" が不一致です`;
          // 先にエラー文言が入っていた場合、書き換え後の文言を記載
          if (WorkData[0][0][26] != undefined) {
            RPA.Logger.info('エラー文言が入っているため追記します');
            const NewErrorText =
              WorkData[0][0][26] +
              `\n\n----------------------------\n\n${ErrorText}`;
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AB${Row[0]}:AB${Row[0]}`,
              values: [[NewErrorText]]
            });
          } else {
            // 番組表下書きシートのAB列にエラー文言を記載
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AB${Row[0]}:AB${Row[0]}`,
              values: [[ErrorText]]
            });
          }
        }
        if (i == 9) {
          RPA.Logger.info(`${Link3}の "曜日" が不一致です`);
          const ErrorText = `【${Link3}】\n\n "曜日" が不一致です`;
          // 先にエラー文言が入っていた場合、書き換え後の文言を記載
          if (WorkData[0][0][26] != undefined) {
            RPA.Logger.info('エラー文言が入っているため追記します');
            const NewErrorText =
              WorkData[0][0][26] +
              `\n\n----------------------------\n\n${ErrorText}`;
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AB${Row[0]}:AB${Row[0]}`,
              values: [[NewErrorText]]
            });
          } else {
            // 番組表下書きシートのAB列にエラー文言を記載
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AB${Row[0]}:AB${Row[0]}`,
              values: [[ErrorText]]
            });
          }
        }
      }
    } else {
      RPA.Logger.info('日付が不一致のため、開始時間で判定します');
      const StartTime = PageDateText2[2].replace(/\s+/g, '');
      const PageDateText3 = StartTime.split(/[:]/);
      const PageDateText4 = OnAirDate.split(/[/]/);
      var m = ('0' + PageDateText4[0]).slice(-2);
      var d = ('0' + PageDateText4[1]).slice(-2);
      var y = WorkData[0][0][1].slice(0, -5);
      var y2 = y.slice(-2);
      var dateAttr = new Date(`${y}-${m}-${d}`);
      var monthAtrr = dateAttr.getMonth();
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
      var mS = month.toString();
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
      // コピーライトを取得
      const CopyRight = await RPA.WebBrowser.findElementByClassName(
        'com-o-Footer__copyright'
      );
      const CopyRightText = await CopyRight.getText();
      RPA.Logger.info(CopyRightText);
      // 開始時間の比較
      const date1 = new Date(
        `${d}-${mS}-${y2} ${PageDateText3[0]}:${PageDateText3[1]}:00 GMT`
      );
      const date2 = new Date(`${d}-${mS}-${y2} 02:00:00 GMT`);
      RPA.Logger.info(date1);
      RPA.Logger.info(date2);
      if (date1 < date2) {
        RPA.Logger.info('開始時間が2:00までなので合っています');
        if (i == 7) {
          const DateSplit = WorkData[0][0][0].split(/[/]/);
          // AC列に月を入力
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AC${Row[0]}:AC${Row[0]}`,
            values: [[DateSplit[0]]]
          });
          // AD列に日を入力
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AD${Row[0]}:AD${Row[0]}`,
            values: [[DateSplit[1]]]
          });
          // AE列に曜日を入力
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AE${Row[0]}:AE${Row[0]}`,
            values: [[w2]]
          });
          // AF列に時間帯を入力
          if (date >= T4_00 && date <= T8_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['朝']]
            });
          }
          if (date >= T9_00 && date <= T11_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['午前']]
            });
          }
          if (date >= T12_00 && date <= T12_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['昼']]
            });
          }
          if (date >= T13_00 && date <= T17_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['午後']]
            });
          }
          if (date >= T18_00 && date <= T23_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['夜']]
            });
          }
          if (date >= T24_00 && date <= T26_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AF${Row[0]}:AF${Row[0]}`,
              values: [['今夜']]
            });
          }
          // AG列に開始時間を入力
          if (Number(PageDateText3[0]) > 12) {
            const ampm = Number(PageDateText3[0]) - 12;
            const ampmStartTime = ampm + ':' + PageDateText3[1];
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AG${Row[0]}:AG${Row[0]}`,
              values: [[`${ampmStartTime}〜`]]
            });
          } else {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AG${Row[0]}:AG${Row[0]}`,
              values: [[`${Number(PageDateText3[0])}:${PageDateText3[1]}〜`]]
            });
          }
          // AO列にコピーライトを入力
          if (CopyRightText.indexOf('Abema') > -1 || CopyRightText.indexOf('abema') > -1 || CopyRightText.indexOf('ABEMA') > -1) {
            RPA.Logger.info('AbemaTVです');
          } else {
            RPA.Logger.info('AbemaTV以外です');
            // 「(C)」もしくは「制作著作」の文字が含まれている場合、「©」に変換して記載
            if (CopyRightText.indexOf('(C)') > -1) {
              const CopyRightTextReplace = CopyRightText.replace('(C)', '©');
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${SSID}`,
                range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                values: [[CopyRightTextReplace]]
              });
            } else if (CopyRightText.indexOf('制作著作') > -1) {
              const CopyRightTextReplace = CopyRightText.replace(
                '制作著作',
                '©'
              );
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${SSID}`,
                range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                values: [[CopyRightTextReplace]]
              });
            } else {
              await RPA.Google.Spreadsheet.setValues({
                spreadsheetId: `${SSID}`,
                range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                values: [[CopyRightText]]
              });
            }
          }
        }
        if (i == 8) {
          // AH列に時間帯を入力
          if (date >= T4_00 && date <= T8_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['朝']]
            });
          }
          if (date >= T9_00 && date <= T11_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['午前']]
            });
          }
          if (date >= T12_00 && date <= T12_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['昼']]
            });
          }
          if (date >= T13_00 && date <= T17_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['午後']]
            });
          }
          if (date >= T18_00 && date <= T23_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['夜']]
            });
          }
          if (date >= T24_00 && date <= T26_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AH${Row[0]}:AH${Row[0]}`,
              values: [['今夜']]
            });
          }
          // AI列に開始時間を入力
          if (Number(PageDateText3[0]) > 12) {
            const ampm = Number(PageDateText3[0]) - 12;
            const ampmStartTime = ampm + ':' + PageDateText3[1];
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AI${Row[0]}:AI${Row[0]}`,
              values: [[`${ampmStartTime}〜`]]
            });
          } else {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AI${Row[0]}:AI${Row[0]}`,
              values: [[`${Number(PageDateText3[0])}:${PageDateText3[1]}〜`]]
            });
          }
          // AO列にコピーライトを入力
          if (CopyRightText.indexOf('Abema') > -1 || CopyRightText.indexOf('abema') > -1 || CopyRightText.indexOf('ABEMA') > -1) {
            RPA.Logger.info('AbemaTVです');
          } else {
            RPA.Logger.info('AbemaTV以外です');
            // 「(C)」もしくは「制作著作」の文字が含まれている場合、「©」に変換して記載
            if (CopyRightText.indexOf('(C)') > -1) {
              const CopyRightTextReplace = CopyRightText.replace('(C)', '©');
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText =
                  WorkData[0][0][39] + ` ${CopyRightTextReplace}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightTextReplace]]
                });
              }
            } else if (CopyRightText.indexOf('制作著作') > -1) {
              const CopyRightTextReplace = CopyRightText.replace(
                '制作著作',
                '©'
              );
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText =
                  WorkData[0][0][39] + ` ${CopyRightTextReplace}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightTextReplace]]
                });
              }
            } else {
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText = WorkData[0][0][36] + ` ${Text}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightText]]
                });
              }
            }
          }
        }
        if (i == 9) {
          // AJ列に時間帯を入力
          if (date >= T4_00 && date <= T8_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['朝']]
            });
          }
          if (date >= T9_00 && date <= T11_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['午前']]
            });
          }
          if (date >= T12_00 && date <= T12_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['昼']]
            });
          }
          if (date >= T13_00 && date <= T17_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['午後']]
            });
          }
          if (date >= T18_00 && date <= T23_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['夜']]
            });
          }
          if (date >= T24_00 && date <= T26_59) {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AJ${Row[0]}:AJ${Row[0]}`,
              values: [['今夜']]
            });
          }
          // AK列に開始時間を入力
          if (Number(PageDateText3[0]) > 12) {
            const ampm = Number(PageDateText3[0]) - 12;
            const ampmStartTime = ampm + ':' + PageDateText3[1];
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AK${Row[0]}:AK${Row[0]}`,
              values: [[`${ampmStartTime}〜`]]
            });
          } else {
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AK${Row[0]}:AK${Row[0]}`,
              values: [[`${Number(PageDateText3[0])}:${PageDateText3[1]}〜`]]
            });
          }
          // AO列にコピーライトを入力
          if (CopyRightText.indexOf('Abema') > -1 || CopyRightText.indexOf('abema') > -1 || CopyRightText.indexOf('ABEMA') > -1) {
            RPA.Logger.info('AbemaTVです');
          } else {
            RPA.Logger.info('AbemaTV以外です');
            // 「(C)」もしくは「制作著作」の文字が含まれている場合、「©」に変換して記載
            if (CopyRightText.indexOf('(C)') > -1) {
              const CopyRightTextReplace = CopyRightText.replace('(C)', '©');
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText =
                  WorkData[0][0][39] + ` ${CopyRightTextReplace}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightTextReplace]]
                });
              }
            } else if (CopyRightText.indexOf('制作著作') > -1) {
              const CopyRightTextReplace = CopyRightText.replace(
                '制作著作',
                '©'
              );
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText =
                  WorkData[0][0][39] + ` ${CopyRightTextReplace}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightTextReplace]]
                });
              }
            } else {
              // 先にコピーライトが入っていた場合、書き換え後の文言を記載
              if (WorkData[0][0][39] != undefined) {
                RPA.Logger.info('コピーライトが入っているため追記します');
                const NewCopyRightText = WorkData[0][0][36] + ` ${Text}`;
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[NewCopyRightText]]
                });
              } else {
                await RPA.Google.Spreadsheet.setValues({
                  spreadsheetId: `${SSID}`,
                  range: `${SSName}!AO${Row[0]}:AO${Row[0]}`,
                  values: [[CopyRightText]]
                });
              }
            }
          }
        }
        RPA.Logger.info('画像をダウンロードします');
        // 画像のURLを取得
        const ImageUrlText = await RPA.WebBrowser.driver
          .findElement(
            By.css(
              '#main > div > div.c-application-DesktopAppContainer__content-container > div.c-application-DesktopAppContainer__content > main > div > div > div > div.c-tv-UpcomingSlotContainerView__header > div.c-tv-UpcomingSlotContainerView__thumbnail-column > div > div > img'
            )
          )
          .getAttribute('src');
        const ImageUrlSplit = ImageUrlText.split('?');
        const ImageUrlRename = ImageUrlSplit[0].replace(/v\d+.webp/, 'jpg');
        RPA.Logger.info(ImageUrlRename);
        const ImageUrlRenameSplit = ImageUrlRename.split('/');
        // 名前を変更
        const Renamming = ImageUrlRenameSplit[6].replace(
          /thumb001.jpg/,
          `${ImageUrlRenameSplit[5]}.jpg`
        );
        RPA.Logger.info(Renamming);
        if (i == 7) {
          // 番組表下書きシートのAL列に画像名を記載
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AL${Row[0]}:AL${Row[0]}`,
            values: [[Renamming]]
          });
        }
        if (i == 8) {
          // 番組表下書きシートのAM列に画像名を記載
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AM${Row[0]}:AM${Row[0]}`,
            values: [[Renamming]]
          });
        }
        if (i == 9) {
          // 番組表下書きシートのAN列に画像名を記載
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AN${Row[0]}:AN${Row[0]}`,
            values: [[Renamming]]
          });
        }
        // ダウンロードフォルダに保存
        request(
          { method: 'GET', url: ImageUrlRename, encoding: null },
          function(error, response, body) {
            if (!error && response.statusCode === 200) {
              fs.writeFileSync(
                `${DownloadFolder}/${Renamming}`,
                body,
                'binary'
              );
            }
          }
        );
        RPA.Logger.info('ダウンロード完了です');
      } else {
        RPA.Logger.info('日付が不一致のため記載ミスです');
        if (i == 7) {
          RPA.Logger.info(`${Link1}の "日付" が不一致です`);
          const ErrorText = `【${Link1}】\n\n "日付" が不一致です`;
          // 番組表下書きシートのAB列にエラー文言を記載
          await RPA.Google.Spreadsheet.setValues({
            spreadsheetId: `${SSID}`,
            range: `${SSName}!AB${Row[0]}:AB${Row[0]}`,
            values: [[ErrorText]]
          });
        }
        if (i == 8) {
          RPA.Logger.info(`${Link2}の "日付" が不一致です`);
          const ErrorText = `【${Link2}】\n\n "日付" が不一致です`;
          // 先にエラー文言が入っていた場合、書き換え後の文言を記載
          if (WorkData[0][0][26] != undefined) {
            RPA.Logger.info('エラー文言が入っているため追記します');
            const NewErrorText =
              WorkData[0][0][26] +
              `\n\n----------------------------\n\n${ErrorText}`;
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AB${Row[0]}:AB${Row[0]}`,
              values: [[NewErrorText]]
            });
          } else {
            // 番組表下書きシートのAB列にエラー文言を記載
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AB${Row[0]}:AB${Row[0]}`,
              values: [[ErrorText]]
            });
          }
        }
        if (i == 9) {
          RPA.Logger.info(`${Link3}の "日付" が不一致です`);
          const ErrorText = `【${Link3}】\n\n "日付" が不一致です`;
          // 先にエラー文言が入っていた場合、書き換え後の文言を記載
          if (WorkData[0][0][26] != undefined) {
            RPA.Logger.info('エラー文言が入っているため追記します');
            const NewErrorText =
              WorkData[0][0][26] +
              `\n\n----------------------------\n\n${ErrorText}`;
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AB${Row[0]}:AB${Row[0]}`,
              values: [[NewErrorText]]
            });
          } else {
            // 番組表下書きシートのAB列にエラー文言を記載
            await RPA.Google.Spreadsheet.setValues({
              spreadsheetId: `${SSID}`,
              range: `${SSName}!AB${Row[0]}:AB${Row[0]}`,
              values: [[ErrorText]]
            });
          }
        }
      }
    }
  }
}

async function PhotoshopData(WorkData, Row) {
  // AP列にPSD保存用の画像名を記載
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!AP${Row[0]}:AP${Row[0]}`,
    values: [[`${WorkData[0][0][1]}.psd`]]
  });
  // Photoshopで使用するデータ(AC〜AP列)を取得
  const PhotoshopData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!AC${Row[0]}:AP${Row[0]}`
  });
  RPA.Logger.info(PhotoshopData);

  await exportCSV(PhotoshopData);

  // 配列をcsvで保存
  async function exportCSV(content) {
    for (var i = 0; i < content.length; i++) {
      var value = content[i];
      for (var j = 0; j < value.length; j++) {
        var innerValue = value[j] === null ? '' : value[j].toString();
        var result = innerValue.replace(/"/g, '""');
        if (result.search(/("|,|\n)/g) >= 0) result = '"' + result + '"';
        if (j > 0) formatCSV += ',';
        formatCSV += result;
      }
      formatCSV += '\n';
    }
    fs.writeFile('formList.csv', formatCSV, 'utf8', function(err) {
      if (err) {
        RPA.Logger.info('保存できませんでした');
      } else {
        RPA.Logger.info('保存しました');
      }
    });
  }
}
