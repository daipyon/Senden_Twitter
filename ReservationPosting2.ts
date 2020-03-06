import RPA from 'ts-rpa';
import { WebDriver, By, FileDetector, Key } from 'selenium-webdriver';
import { rootCertificates } from 'tls';
import { worker } from 'cluster';
import { cachedDataVersionTag } from 'v8';
import { start } from 'repl';
import { Command } from 'selenium-webdriver/lib/command';
import { Driver } from 'selenium-webdriver/safari';
const fs = require('fs');
// デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
RPA.Logger.level = 'INFO';

// スプレッドシートIDとシート名を記載
const SSID = process.env.Senden_Twitter_SheetID;
const SSName = process.env.Senden_Twitter_SheetName;
// アイパス一覧シートIDとシート名を記載
const SSID2 = process.env.Senden_Twitter_SheetID2;
const SSName2 = process.env.Senden_Twitter_SheetName2;
// 画像などを保存するフォルダのパスを記載 ※.envファイルは同じにしない
const DownloadFolder = process.env.Senden_Twitter_RPAFolder2;
// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;

const FilePathData = [''];
const FirstLoginFlag = ['true'];
async function WorkStart() {
  await RPA.Google.authorize({
    // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
  });

  // 作業対象行とデータを取得
  const FilePathDate = [''];
  const WorkData = [];
  const Row = [];
  await GetDataRow(FilePathData, FilePathDate, WorkData, Row);

  // Twitterにログイン
  if (FirstLoginFlag[0] == 'true') {
    await TwitterLogin();
  }
  if (FirstLoginFlag[0] == 'false') {
    await TwitterLogin2();
  }
  // 一度ログインしたら、以降はログインページをスキップ
  FirstLoginFlag[0] = 'false';

  // ウェブサイトカードを作成
  await CreateWebSiteCard(FilePathData, FilePathDate);

  // Twitter投稿を予約
  await ReservationTweet(WorkData, FilePathData, FilePathDate);

  // 作業用フォルダを空にするためファイルを削除
  await DeleteFile();
}

async function Start() {
  try {
    await RPA.Google.authorize({
      //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });
    // ダウンロードフォルダから画像のパスを取得
    const path = require('path');
    const dirPath = path.resolve(DownloadFolder);
    const FirstData = [];
    FirstData[0] = fs.readdirSync(dirPath);
    for (let i in FirstData[0]) {
      if (FirstData[0][i].indexOf('.png') > 0) {
        FilePathData[0] = FirstData[0][i];
        RPA.Logger.info(FilePathData);
        RPA.Logger.info('作業を開始します');
        await WorkStart();
      }
    }
  } catch (error) {
    RPA.SystemLogger.error(error);
    await RPA.WebBrowser.takeScreenshot();
    // Slackにも通知
    await RPA.Slack.chat.postMessage({
      token: Slack_Token,
      channel: Slack_Channel,
      text: '【宣伝_Twitter 予約投稿（番組表）】（CAPC-1375用）でエラーが発生しました！'
    });
  }
  RPA.Logger.info('作業を終了します');
  await RPA.WebBrowser.quit();
}

Start();

async function GetDataRow(FilePathData, FilePathDate, WorkData, Row) {
  // ".png"の文字を削除
  FilePathDate[0] = FilePathData[0].slice(0, -4);
  RPA.Logger.info(FilePathDate);
  const JudgeData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!C6:C30000`
  });
  for (let i in JudgeData) {
    if (FilePathDate[0] == JudgeData[i][0]) {
      RPA.Logger.info(JudgeData[i][0]);
      Row[0] = Number(i) + 6;
      break;
    }
  }
  RPA.Logger.info('この行のデータを取得します → ', Row[0]);
  // 番組表下書きシートのデータ(B〜AA列)を取得
  WorkData[0] = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!B${Row[0]}:AA${Row[0]}`
  });
  RPA.Logger.info(WorkData[0]);
}

async function TwitterLogin() {
  // アイパス一覧シートからTwitterのIDとパスワードを取得
  const TwitterID = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID2}`,
    // spreadsheetId: `${SSID}`,
    range: `${SSName2}!D3:D3`
  });
  const TwitterPW = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID2}`,
    // spreadsheetId: `${SSID}`,
    range: `${SSName2}!E3:E3`
  });
  // Twitterにログイン
  await RPA.WebBrowser.get(process.env.Twitter_Login_URL);
  await RPA.sleep(2000);
  try {
    // ログイン画面の仕様が普段と異なる際の処理
    var LoginFlag = ['false'];
    const Login = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="react-root"]/div/div/div/header/div[2]/div[1]/div/div[2]/div[1]/div[1]/a'
    );
    await RPA.WebBrowser.mouseClick(Login);
    await RPA.sleep(1000);
    const TwitterLoginId = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="react-root"]/div/div/div[1]/main/div/div/form/div/div[1]/label/div[2]/div/input'
    );
    await RPA.WebBrowser.sendKeys(TwitterLoginId, [TwitterID[0][0]]);
    await RPA.sleep(500);
    const TwitterLoginPw = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="react-root"]/div/div/div[1]/main/div/div/form/div/div[2]/label/div[2]/div/input'
    );
    await RPA.WebBrowser.sendKeys(TwitterLoginPw, [TwitterPW[0][0]]);
    const LoginButton = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="react-root"]/div/div/div[1]/main/div/div/form/div/div[3]/div'
    );
    await RPA.WebBrowser.mouseClick(LoginButton);
    await RPA.sleep(3000);
  } catch {
    LoginFlag[0] = 'true';
    const TwitterLoginId = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="react-root"]/div/div/div[1]/main/div/div/form/div/div[1]/label/div/div[2]/div/input'
    );
    await RPA.WebBrowser.sendKeys(TwitterLoginId, [TwitterID[0][0]]);
    await RPA.sleep(500);
    const TwitterLoginPw = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="react-root"]/div/div/div[1]/main/div/div/form/div/div[2]/label/div/div[2]/div/input'
    );
    await RPA.WebBrowser.sendKeys(TwitterLoginPw, [TwitterPW[0][0]]);
    const LoginButton = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="react-root"]/div/div/div[1]/main/div/div/form/div/div[3]/div'
    );
    await RPA.WebBrowser.mouseClick(LoginButton);
    await RPA.sleep(3000);
  }
  // 【もっと見る】をクリック
  const Menu = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="react-root"]/div/div/div/header/div/div/div/div/div[2]/nav/div/div/div'
  );
  await RPA.WebBrowser.mouseClick(Menu);
  await RPA.sleep(500);
  // 【Twitter広告】をクリック
  const TwitterAd = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="react-root"]/div/div/div[1]/div/div/div[2]/div[3]/div/div/div/div/div[5]/a'
  );
  await RPA.WebBrowser.mouseClick(TwitterAd);
  await RPA.sleep(3000);
  // ウィンドウの切り替え
  const windows = await RPA.WebBrowser.getAllWindowHandles();
  await RPA.WebBrowser.switchToWindow(windows[1]);
  // 広告に利用するアカウントを選択
  const AdAccount = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="account-selector-form"]/ul/li[1]/div'
  );
  await RPA.WebBrowser.mouseClick(AdAccount);
  await RPA.sleep(3000);
  // 【クリエイティブ】をクリック
  const Creative = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="SharedNavBarContainer"]/div/div/ul[1]/li[2]/a'
  );
  await RPA.WebBrowser.mouseClick(Creative);
  // 【カード】をクリック
  const Card = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="SharedNavBarContainer"]/div/div/ul[1]/li[2]/ul/li[2]/a'
  );
  await RPA.WebBrowser.mouseClick(Card);
  await RPA.sleep(5000);
}

async function TwitterLogin2() {
  RPA.Logger.info('カードライブラリページに直接遷移します');
  await RPA.WebBrowser.get(process.env.Twitter_CardLibrary_URL);
  await RPA.sleep(5000);
}

async function CreateWebSiteCard(FilePathData, FilePathDate) {
  // 【カードを作成】をクリック
  const CreateCard = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="root"]/div/div[1]/div[1]/div/button'
  );
  await RPA.WebBrowser.mouseClick(CreateCard);
  // 【ウェブサイトカード】をクリック
  const WebSiteCard = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="feather-dropdown-0-menu-item-content-0"]'
  );
  await RPA.WebBrowser.mouseClick(WebSiteCard);
  // 【画像を選択】をクリック
  const SelectImage = await RPA.WebBrowser.findElementByXPath(
    '/html/body/div[5]/div/div/div[2]/div[3]/form/div[1]/button'
  );
  await RPA.WebBrowser.mouseClick(SelectImage);
  // InputFile に直接パスを打ち込んでアップロード
  const InputFile = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      xpath: '/html/body/div[5]/div/div/div[2]/div[3]/div/div[2]/div[1]/input'
    }),
    8000
  );
  await RPA.WebBrowser.sendKeys(InputFile, [
    `${DownloadFolder}/${FilePathData[0]}`
  ]);
  await RPA.sleep(3000);
  // 【ヘッドライン】を入力
  const HeadLine = await RPA.WebBrowser.findElementByXPath(
    '/html/body/div[5]/div/div/div[2]/div[3]/form/label[1]/input'
  );
  await RPA.WebBrowser.sendKeys(HeadLine, ['△タップして番組表一覧を見る']);
  // 【ウェブサイトのURL】を入力
  const WebSiteUrl = await RPA.WebBrowser.findElementByXPath(
    '/html/body/div[5]/div/div/div[2]/div[3]/form/label[2]/input'
  );
  await RPA.WebBrowser.sendKeys(WebSiteUrl, [process.env.Twitter_WebSite_URL]);
  // 【カード名】を入力
  const CardName = await RPA.WebBrowser.findElementByXPath(
    '/html/body/div[5]/div/div/div[2]/div[3]/form/label[3]/input'
  );
  await RPA.WebBrowser.sendKeys(CardName, [FilePathDate[0]]);
  // await RPA.WebBrowser.sendKeys(CardName, [`※※※RPAテストです※※※ ${FilePathDate[0]}`]);
  // 【作成】をクリック
  const Create = await RPA.WebBrowser.findElementByXPath(
    '/html/body/div[5]/div/div/div[3]/div/button[3]'
  );
  await RPA.WebBrowser.mouseClick(Create);
  await RPA.sleep(5000);
}

async function ReservationTweet(WorkData, FilePathData, FilePathDate) {
  // 読み込みエラーが発生した場合はブラウザを更新
  while (0 == 0) {
    try {
      var LoadingFlag = ['false'];
      const Loading = await RPA.WebBrowser.findElementByXPath(
        '//*[@id="root"]/div/div[1]/div[2]/div/div[2]/div[1]/h3'
      );
      const LoadingText = await Loading.getText();
      if (LoadingText == 'このコンテンツの読み込み中に問題が発生しました。') {
        RPA.Logger.info('読み込みエラーが発生したため、ブラウザを更新します');
        await RPA.WebBrowser.refresh();
        await RPA.sleep(5000);
      }
    } catch {
      LoadingFlag[0] = 'true';
      RPA.Logger.info('読み込みエラーは発生しませんでした。');
      break;
    }
  }
  // 再度読み込みエラーが発生した場合は、ブラウザを更新し再検索
  while (0 == 0) {
    // カード名を検索
    for (var i = 1; i <= 10; i++) {
      const SearchCardName = await RPA.WebBrowser.findElementByXPath(
        `//*[@id="root"]/div/div[1]/div[2]/div/div[2]/div/div[1]/div[${Number(
          i
        )}]/div[2]/div[1]`
      );
      const SearchCardNameText = await SearchCardName.getText();
      RPA.Logger.info(SearchCardNameText);
      if (SearchCardNameText == FilePathDate[0]) {
        // if (SearchCardNameText == `※※※RPAテストです※※※ ${FilePathDate[0]}`) {
        // 作成したカードにマウスオーバー
        RPA.Logger.info('    ↑     一致しました');
        const CreatedCard = await RPA.WebBrowser.findElementByXPath(
          `//*[@id="root"]/div/div[1]/div[2]/div/div[2]/div/div[1]/div[${Number(
            i
          )}]/div[1]/div[2]/div/a[2]`
        );
        await RPA.WebBrowser.mouseMove(CreatedCard);
        // 【ツイート】をクリック
        await RPA.WebBrowser.mouseClick(CreatedCard);
        break;
      }
    }
    try {
      var LoadingFlag2 = ['false'];
      const Loading = await RPA.WebBrowser.findElementByXPath(
        '//*[@id="root"]/div/div[1]/div[2]/div/div[2]/div[1]/h3'
      );
      const LoadingText = await Loading.getText();
      if (LoadingText == 'このコンテンツの読み込み中に問題が発生しました。') {
        RPA.Logger.info('読み込みエラーが発生したため、ブラウザを更新します');
        await RPA.WebBrowser.refresh();
        await RPA.sleep(5000);
      }
    } catch {
      LoadingFlag2[0] = 'true';
      RPA.Logger.info('読み込みエラーは発生しませんでした。');
      break;
    }
  }
  await RPA.sleep(10000);
  // ウィンドウの切り替え
  const windows = await RPA.WebBrowser.getAllWindowHandles();
  await RPA.WebBrowser.switchToWindow(windows[2]);
  // 番組表下書きシートのM列の文言を入力
  RPA.Logger.info(WorkData[0][0][11]);
  const PostWording = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="root"]/div/div[1]/div/div/div/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div[1]/textarea'
  );
  const Text1 = await encodeURI(WorkData[0][0][11]);
  await RPA.WebBrowser.driver.executeScript(
    `document.getElementsByClassName("KYqgv_Qj--TweetComposer-tweetTextarea")[0].value = decodeURI("${Text1}")`
  );
  await RPA.WebBrowser.sendKeys(PostWording, [` `]);
  await RPA.WebBrowser.scrollTo({
    xpath:
      '//*[@id="root"]/div/div[1]/div/div/div/div/div/div[2]/div/div/div[2]'
  });
  // 【広告用】のチェックを外す
  const Uncheck = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="root"]/div/div[1]/div/div/div/div/div/div[2]/div/div/div[4]/div[1]/div[1]/label/input'
  );
  await RPA.WebBrowser.mouseClick(Uncheck);
  // カーソルをクリック
  const Cursor = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="root"]/div/div[1]/div/div/div/div/div/div[2]/div/div/div[4]/div[1]/div[2]/button[2]'
  );
  await RPA.WebBrowser.mouseClick(Cursor);
  await RPA.sleep(1000);
  // 【予約設定する】をクリック
  const ReservationSettings = await RPA.WebBrowser.findElementByXPath(
    '/html/body/div[4]/div/ul/li[1]/button'
  );
  await RPA.WebBrowser.mouseClick(ReservationSettings);
  // カレンダーをクリック
  const Calender = await RPA.WebBrowser.findElementByXPath(
    '/html/body/div[2]/div/div[1]/div/div/div/div/div/div[2]/div/div/div[4]/div/button/span[2]'
  );
  await RPA.WebBrowser.mouseClick(Calender);
  // カレンダーの月を取得
  const MonthButton = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="feather-dropdown-7"]/ul/li/div/div[1]/fieldset[1]/button/span[1]'
  );
  const MonthText = await MonthButton.getText();
  // カレンダーの"月"という文字を削除
  const Month = MonthText.slice(0, -1);
  RPA.Logger.info(Month);
  // ファイル名の月の部分を取得
  const FilePath_Month = FilePathData[0].slice(5, -6);
  RPA.Logger.info(FilePath_Month);
  if (Number(Month) != Number(FilePath_Month)) {
    RPA.Logger.info('月が不一致のため、次月を指定します');
    const NextMonthButton = await RPA.WebBrowser.findElementByXPath(
      '//*[@id="feather-dropdown-7"]/ul/li/div/div[1]/button[2]'
    );
    await RPA.WebBrowser.mouseClick(NextMonthButton);
  }
  if (Number(Month) == Number(FilePath_Month)) {
    RPA.Logger.info('月が一致しました');
  }
  // ファイル名の日付の部分を取得
  const FilePath_Day = FilePathData[0].slice(7, -4);
  RPA.Logger.info(FilePath_Day);
  // 1日の位置を検索
  const ThroughDay = [];
  for (var i = 1; i <= 7; i++) {
    const SearchDay = await RPA.WebBrowser.findElementByXPath(
      `//*[@id="feather-dropdown-7"]/ul/li/div/div[2]/div[2]/div[1]/div[${Number(
        i
      )}]/span`
    );
    const SearchDayText = await SearchDay.getText();
    if (SearchDayText != '1') {
      RPA.Logger.info(`${SearchDayText}日`);
      RPA.Logger.info('1日ではないためスルーします');
      ThroughDay.push(SearchDayText);
    } else if (SearchDayText == '1') {
      RPA.Logger.info(`${SearchDayText}日`);
      RPA.Logger.info('1日です');
      RPA.Logger.info('判定処理を開始します');
      break;
    }
  }
  // 1日から判定開始
  const JudgeDayFlag = [];
  JudgeDayFlag[0] = false;
  for (var i = ThroughDay.length + 1; i <= 7; i++) {
    const SearchDay = await RPA.WebBrowser.findElementByXPath(
      `//*[@id="feather-dropdown-7"]/ul/li/div/div[2]/div[2]/div[1]/div[${Number(
        i
      )}]/span`
    );
    const SearchDayText = await SearchDay.getText();
    // 日付にマウスオーバー
    const Day_MouseOver = await RPA.WebBrowser.findElementByXPath(
      `//*[@id="feather-dropdown-7"]/ul/li/div/div[2]/div[2]/div[1]/div[${Number(
        i
      )}]`
    );
    await RPA.WebBrowser.mouseMove(Day_MouseOver);
    await RPA.sleep(50);
    if (Number(SearchDayText) != Number(FilePath_Day)) {
      RPA.Logger.info('日付が不一致です');
    } else if (Number(SearchDayText) == Number(FilePath_Day)) {
      JudgeDayFlag[0] = true;
      RPA.Logger.info('日付が一致しました');
      // 一致した日付をクリック
      await RPA.WebBrowser.mouseClick(Day_MouseOver);
      break;
    }
  }
  if (JudgeDayFlag[0] == false) {
    RPA.Logger.info('次の処理に進みます');
    await JudgeDay(FilePath_Day);
  }

  // カレンダーの2列目から判定再開
  async function JudgeDay(FilePath_Day) {
    const JudgeDayFlag = [];
    JudgeDayFlag[0] = false;
    for (var i = 2; i <= 6; i++) {
      for (var n = 1; n <= 7; n++) {
        const SearchDay = await RPA.WebBrowser.findElementByXPath(
          `//*[@id="feather-dropdown-7"]/ul/li/div/div[2]/div[2]/div[${Number(
            i
          )}]/div[${Number(n)}]/span`
        );
        const SearchDayText = await SearchDay.getText();
        // 日付にマウスオーバー
        const Day_MouseOver = await RPA.WebBrowser.findElementByXPath(
          `//*[@id="feather-dropdown-7"]/ul/li/div/div[2]/div[2]/div[${Number(
            i
          )}]/div[${Number(n)}]`
        );
        await RPA.WebBrowser.mouseMove(Day_MouseOver);
        await RPA.sleep(50);
        if (Number(SearchDayText) == Number(FilePath_Day)) {
          JudgeDayFlag[0] = true;
          RPA.Logger.info('日付が一致しました');
          // 一致した日付をクリック
          await RPA.WebBrowser.mouseClick(Day_MouseOver);
          break;
        }
      }
      if (JudgeDayFlag[0] == true) {
        RPA.Logger.info('判定処理を終了します');
        break;
      }
    }
  }

  // 予約時間を入力
  const ReservationTime = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="feather-dropdown-7"]/ul/li/div/div[3]/input'
  );
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await RPA.WebBrowser.sendKeys(ReservationTime, ['8:00am']);
  await ReservationTime.sendKeys(Key.ENTER);
  // 【予約設定する】をクリック
  const ReservationSettings2 = await RPA.WebBrowser.findElementByXPath(
    '//*[@id="root"]/div/div[1]/div/div/div/div/div/div[2]/div/div/div[4]/div/div[2]/button[1]'
  );
  await RPA.WebBrowser.mouseClick(ReservationSettings2);
  await RPA.sleep(5000);
  // 現在のウィンドウを閉じる
  await RPA.WebBrowser.driver.close();
  await RPA.WebBrowser.switchToWindow(windows[1]);
}

// 取得したファイルを削除
async function DeleteFile() {
  const FirstData = fs.readdirSync(DownloadFolder);
  RPA.Logger.info(FirstData);
  RPA.Logger.info('予約が完了しましたので、ファイルを削除します');
  for (let i in FirstData) {
    if (FirstData[i].indexOf('.png') > 0) {
      fs.unlink(`${DownloadFolder}/${FirstData[i]}`, function(err) {});
      break;
    }
  }
}
