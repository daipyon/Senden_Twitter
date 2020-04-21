import RPA from 'ts-rpa';
import { Key, WebElement } from 'selenium-webdriver';
const fs = require('fs');

// SlackのトークンとチャンネルID
const Slack_Token = process.env.AbemaTV_RPAError_Token;
const Slack_Channel = process.env.AbemaTV_RPAError_Channel;
const Slack_Text = [`【Twitter 予約投稿】予約完了しました`];

// スプレッドシートIDとシート名を記載
// const mySSID = process.env.My_SheetID;
const SSID = process.env.Senden_Twitter_SheetID;
const SSName = process.env.Senden_Twitter_SheetName;
// アイパス一覧シート
const SSID2 = process.env.Senden_Twitter_SheetID2;
const SSName2 = process.env.Senden_Twitter_SheetName2;
// 作業するスプレッドシートから読み込む行数を記載
const StartRow = 6;
const LastRow = 30000;

// 画像などを保存するフォルダのパスを記載
const DLFolder = __dirname + '/DL/';

// 作業対象行とデータを取得
const WorkData = [];
const Row = [];

// エラー発生時のテキストを格納
const ErrorText = [];

async function Start() {
  if (ErrorText.length == 0) {
    // デバッグログを最小限(INFOのみ)にする ※[DEBUG]が非表示になる
    RPA.Logger.level = 'INFO';
    await RPA.Google.authorize({
      // accessToken: process.env.GOOGLE_ACCESS_TOKEN,
      refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
      tokenType: 'Bearer',
      expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10)
    });

    // テスト用
    // DLフォルダ内のファイルを削除
    // await DeleteFiles();

    // 番組表フォルダからファイルを取得
    const GoogleDriveFolderList = await RPA.Google.Drive.listFiles({
      parents: [`${process.env.Senden_Twitter_GoogleDriveFolderID}`]
    });
    await RPA.Logger.info(GoogleDriveFolderList);
    await RPA.Logger.info(
      `Drive内のファイル個数    → ${GoogleDriveFolderList.length}`
    );
    for (let i in GoogleDriveFolderList) {
      // Google Drive内のファイルをDLフォルダにダウンロード
      await RPA.Google.Drive.download({
        fileId: `${GoogleDriveFolderList[i].id}`
      });
      await Work();
    }
  }
  // エラー発生時の処理
  if (ErrorText.length >= 1) {
    // const DOM = await RPA.WebBrowser.driver.getPageSource();
    // await await RPA.Logger.info(DOM);
    await RPA.SystemLogger.error(ErrorText);
    Slack_Text[0] = `【Twitter 予約投稿】でエラー発生しました\n${ErrorText}`;
    await RPA.WebBrowser.takeScreenshot();
  }
  await RPA.Logger.info('作業を終了します');
  await SlackPost(Slack_Text[0]);
  await RPA.WebBrowser.quit();
  await RPA.sleep(1000);
  await process.exit();
}

Start();

const FirstLoginFlag = ['true'];
async function Work() {
  try {
    // ファイルのパスを保持する関数
    const FilePathData = [];
    // フォルダから動画のパスを取得
    await FilePathGet(FilePathData);
    // ファイルの日付を保持する関数
    const FilePathDate = [];
    // シートのデータを取得
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
    // 作成したカード名を検索
    await SearchCardName(FilePathDate);
    // Twitter投稿を予約
    await ReservationTweet(WorkData, FilePathData);
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

// テスト用
async function DeleteFiles() {
  const ImageData = await fs.readdirSync(DLFolder);
  await RPA.Logger.info(ImageData);
  for (let i in ImageData) {
    await RPA.Logger.info(ImageData[i]);
    if (ImageData[i].indexOf('.png') > 0) {
      await RPA.Logger.info(' ↑ .png ですので削除します ');
      await fs.unlink(`${DLFolder}/${ImageData[i]}`, function(err) {});
    }
  }
  await RPA.Logger.info('ファイルの削除が完了しました');
}

async function FilePathGet(FilePathData) {
  const DLFolderList = await fs.readdirSync(DLFolder);
  await RPA.Logger.info('DLフォルダの一覧　　　   → ' + DLFolderList);
  // .png が含まれているファイルのみ抽出
  for (let i in DLFolderList) {
    await RPA.Logger.info(DLFolderList[i]);
    if (DLFolderList[i].indexOf('.png') > 0) {
      await RPA.Logger.info(' ↑ .png ですので取得します ');
      FilePathData[0] = await DLFolderList[i];
      break;
    }
  }
  await RPA.Logger.info('取得したファイルのパス   → ' + FilePathData);
}

async function GetDataRow(FilePathData, FilePathDate, WorkData, Row) {
  // ".png"の文字を削除
  FilePathDate[0] = await FilePathData[0].slice(0, -4);
  const JudgeData = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!C${StartRow}:C${LastRow}`
  });
  for (let i in JudgeData) {
    if (FilePathDate[0] == JudgeData[i][0]) {
      await RPA.Logger.info('取得したファイルの日付   → ' + FilePathDate[0]);
      await RPA.Logger.info('シートに記載している日付 → ' + JudgeData[i][0]);
      Row[0] = Number(i) + 6;
      break;
    }
  }
  await RPA.Logger.info('この行のデータを取得します → ', Row[0]);
  // 番組表下書きシートのデータ(B〜AA列)を取得
  WorkData[0] = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName}!B${Row[0]}:AA${Row[0]}`
  });
  await RPA.Logger.info(WorkData[0]);
}

async function TwitterLogin() {
  // テスト用
  // throw new Error();

  // アイパス一覧シートからTwitterのIDとパスワードを取得
  const TwitterID = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName2}!D3:D3`
  });
  const TwitterPW = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID2}`,
    range: `${SSName2}!E3:E3`
  });

  // テスト用
  // await RPA.Logger.info(TwitterID[0][0]);
  // await RPA.Logger.info(TwitterPW[0][0]);

  // Twitterにログイン
  await RPA.WebBrowser.get(process.env.Twitter_Login_URL);
  await RPA.sleep(5000);
  const Login: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('css-1dbjc4n r-1awozwy r-1pz39u2 r-18u37iz r-16y2uox')[0].children[0].children[0]`
  );
  await RPA.WebBrowser.mouseClick(Login);
  await RPA.sleep(1000);
  const TwitterLoginId = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      name: 'session[username_or_email]'
    }),
    5000
  );
  await RPA.WebBrowser.sendKeys(TwitterLoginId, [TwitterID[0][0]]);
  await RPA.sleep(500);
  const TwitterLoginPw = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({ name: 'session[password]' }),
    5000
  );
  await RPA.WebBrowser.sendKeys(TwitterLoginPw, [TwitterPW[0][0]]);
  const LoginButton: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('css-1dbjc4n')[21].children[0]`
  );
  await RPA.sleep(500);
  await RPA.WebBrowser.mouseClick(LoginButton);
  while (0 == 0) {
    await RPA.sleep(5000);
    const AbeaniTweet = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({ id: 'accessible-list-0' }),
      5000
    );
    const AbeaniTweetText = await AbeaniTweet.getText();
    if (AbeaniTweetText.length >= 1) {
      await RPA.Logger.info('＊＊＊ログイン成功しました＊＊＊');
      break;
    }
  }
  await RPA.Logger.info('カードライブラリページに直接遷移します');
  await RPA.WebBrowser.get(process.env.Twitter_CardLibrary_URL);
  await RPA.sleep(5000);
}

async function TwitterLogin2() {
  await RPA.Logger.info('カードライブラリページに直接遷移します');
  await RPA.WebBrowser.get(process.env.Twitter_CardLibrary_URL);
  await RPA.sleep(5000);
}

async function CreateWebSiteCard(FilePathData, FilePathDate) {
  // 【カードを作成】をクリック
  const CreateCard = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'Button Button--primary'
    }),
    5000
  );
  await RPA.WebBrowser.mouseClick(CreateCard);
  // 【ウェブサイトカード】をクリック
  const WebSiteCard = await RPA.WebBrowser.findElementById(
    'feather-dropdown-0-menu-item-content-0'
  );
  await RPA.WebBrowser.mouseClick(WebSiteCard);
  // 【画像を選択】をクリック
  const SelectImage: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('FormField')[0].children[1]`
  );
  await RPA.WebBrowser.mouseClick(SelectImage);
  // InputFile に直接パスを打ち込んでアップロード
  // const InputFile = await RPA.WebBrowser.wait(
  //   RPA.WebBrowser.Until.elementLocated({
  //     xpath: '/html/body/div[5]/div/div/div[2]/div[3]/div/div[2]/div[1]/input'
  //   }),
  //   5000
  // );
  await RPA.sleep(1000);
  const InputFile: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByTagName('input')[1]`
  );
  await RPA.WebBrowser.sendKeys(InputFile, [`${DLFolder}${FilePathData[0]}`]);
  await RPA.sleep(3000);
  // 【ヘッドライン】を入力
  const HeadLine: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('FormInput')[1]`
  );
  await RPA.WebBrowser.sendKeys(HeadLine, ['△タップして番組表一覧を見る']);
  // 【ウェブサイトのURL】を入力
  const WebSiteUrl: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('FormInput')[2]`
  );
  await RPA.WebBrowser.sendKeys(WebSiteUrl, [process.env.Twitter_WebSite_URL]);
  // 【カード名】を入力
  const CardName: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('FormInput')[3]`
  );

  // テスト用
  // await RPA.WebBrowser.sendKeys(CardName, [
  //   `※※※RPAテストです※※※ ${FilePathDate[0]}`
  // ]);

  // 本番用
  await RPA.WebBrowser.sendKeys(CardName, [FilePathDate[0]]);

  // 【作成】をクリック
  // const Create: WebElement = await RPA.WebBrowser.driver.executeScript(
  //   `return document.getElementsByClassName('Button Button--primary Button--small')[0]`
  // );
  const Create = await RPA.WebBrowser.findElementByClassName(
    'Button Button--primary Button--small'
  );
  await RPA.WebBrowser.mouseClick(Create);
  await RPA.sleep(5000);
  await RPA.Logger.info('カードを作成しました');
}

const LoadingFlag = ['false'];
async function SearchCardName(FilePathDate) {
  while (0 == 0) {
    try {
      // 読み込みエラーが発生した場合はブラウザを更新
      await BrowserReload();
    } catch {
      LoadingFlag[0] = 'true';
      await RPA.Logger.info('読み込みエラーは発生しませんでした。');
      break;
    }
  }
  // 再度読み込みエラーが発生した場合は、ブラウザを更新し再検索する
  while (0 == 0) {
    // カード名を検索
    for (let i = 0; i <= 9; i++) {
      const SearchCardName: WebElement = await RPA.WebBrowser.driver.executeScript(
        `return document.getElementsByClassName('src-cardsmanager-views-CardsGrid-styles-module--cardName')[${Number(
          i
        )}]`
      );
      const SearchCardNameText = await SearchCardName.getText();
      await RPA.Logger.info(SearchCardNameText);

      // テスト用
      // if (SearchCardNameText == `※※※RPAテストです※※※ ${FilePathDate[0]}`) {

      // 本番用
      if (SearchCardNameText == FilePathDate[0]) {
        await RPA.Logger.info('    ↑     一致しました');
        const CreatedCard: WebElement = await RPA.WebBrowser.driver.executeScript(
          `return document.getElementsByClassName('src-cardsmanager-views-CardsGrid-styles-module--cardOverlay')[${Number(
            i
          )}].children[0].children[1]`
        );
        // 【ツイート】をクリック
        await RPA.WebBrowser.mouseClick(CreatedCard);
        break;
      }
    }
    try {
      LoadingFlag[0] = 'false';
      await BrowserReload();
    } catch {
      LoadingFlag[0] = 'true';
      await RPA.Logger.info('読み込みエラーは発生しませんでした。');
      break;
    }
  }
  await RPA.sleep(5000);
}

async function ReservationTweet(WorkData, FilePathData) {
  // ウィンドウの切り替え
  const windows = await RPA.WebBrowser.getAllWindowHandles();
  await RPA.WebBrowser.switchToWindow(windows[1]);
  // M列（文言）を入力
  const PostWording = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'KYqgv_Qj--TweetComposer-tweetTextarea'
    }),
    10000
  );
  // そのままでは絵文字が入力できないためエンコードする
  const Text1 = await encodeURI(WorkData[0][0][11]);
  await RPA.WebBrowser.driver.executeScript(
    `document.getElementsByClassName("KYqgv_Qj--TweetComposer-tweetTextarea")[0].value = decodeURI("${Text1}")`
  );
  await RPA.WebBrowser.sendKeys(PostWording, [` `]);
  await RPA.WebBrowser.scrollTo({
    selector:
      '#root > div > div.src-tweetcomposer-components-PageLayout-styles-module--root > div > div > div > div > div > div.src-tweetcomposer-components-TweetComposerContainer-styles-module--tweetComposerContainer > div > div > div._1hGMbSMU--TweetComposer-actionsContainer'
  });
  // 【広告用】のチェックを外す
  // const Uncheck: WebElement = await RPA.WebBrowser.driver.executeScript(
  //   `return document.getElementsByClassName('_227h4M8t--FanoutSwitch-checkbox')[0]`
  // );
  const Uncheck = await RPA.WebBrowser.findElementByClassName(
    '_227h4M8t--FanoutSwitch-checkbox'
  );
  await RPA.WebBrowser.mouseClick(Uncheck);
  // カーソルをクリック
  const Cursor: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('ButtonGroup ButtonGroup--primary w0xuLYbb--TweetComposer-submitTweetButton')[0].children[1]`
  );
  await RPA.WebBrowser.mouseClick(Cursor);
  // 【予約設定する】をクリック
  const ReservationSettings: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('Dropdown-menuItemList')[0].children[0].children[0]`
  );
  await RPA.WebBrowser.mouseClick(ReservationSettings);
  // カレンダーをクリック
  // const Calender: WebElement = await RPA.WebBrowser.driver.executeScript(
  //   `return document.getElementsByClassName('Button _1xl1L2Qp--TweetComposer-actionButton DatePickerDropdownTarget')[0]`
  // );
  const Calender = await RPA.WebBrowser.findElementByClassName(
    'Button _1xl1L2Qp--TweetComposer-actionButton DatePickerDropdownTarget'
  );
  await RPA.WebBrowser.mouseClick(Calender);
  // カレンダーでの処理
  await Calendar(FilePathData);
  // 予約時間を入力
  // const ReservationTime: WebElement = await RPA.WebBrowser.driver.executeScript(
  //   `return document.getElementsByClassName('TimePicker FormInput FormInput--small')[0]`
  // );
  const ReservationTime = await RPA.WebBrowser.findElementByClassName(
    'TimePicker FormInput FormInput--small'
  );
  // クリアが効かないため、バックスペースでデフォルトの時間を削除
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await ReservationTime.sendKeys(Key.BACK_SPACE);
  await RPA.sleep(100);
  await RPA.WebBrowser.sendKeys(ReservationTime, ['8:00am']);
  await RPA.sleep(300);
  await ReservationTime.sendKeys(Key.ENTER);
  await RPA.sleep(500);
  // 【予約設定する】をクリック
  const ReservationSettings2: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('ButtonGroup ButtonGroup--primary w0xuLYbb--TweetComposer-submitTweetButton')[0].children[0]`
  );
  await RPA.WebBrowser.mouseClick(ReservationSettings2);
  await RPA.sleep(5000);
  await RPA.Logger.info('予約設定しました');
  // 現在のウィンドウを閉じる
  await RPA.WebBrowser.driver.close();
  await RPA.WebBrowser.switchToWindow(windows[0]);
}

// スルーする日付を保持する変数
const ThroughDay = [];
async function Calendar(FilePathData) {
  // カレンダーの月を取得
  // const MonthButton: WebElement = await RPA.WebBrowser.driver.executeScript(
  //   `return document.getElementsByClassName('Button Button--tertiary Button--xsmall')[0]`
  // );
  const MonthButton = await RPA.WebBrowser.findElementByClassName(
    'Button Button--tertiary Button--xsmall'
  );
  const MonthText = await MonthButton.getText();
  // カレンダーの"月"という文字を削除
  const Month = await MonthText.slice(0, -1);
  await RPA.Logger.info(Month);
  // ファイル名の月の部分を取得
  const FilePath_Month = await FilePathData[0].slice(5, -6);
  await RPA.Logger.info(FilePath_Month);
  if (Number(Month) != Number(FilePath_Month)) {
    await RPA.Logger.info('月が不一致のため、次月を指定します');
    // const NextMonthButton: WebElement = await RPA.WebBrowser.driver.executeScript(
    //   `return document.getElementsByClassName('CalendarNavigation-monthButton CalendarNavigation-monthButton--next')[0]`
    // );
    const NextMonthButton = await RPA.WebBrowser.findElementByClassName(
      'CalendarNavigation-monthButton CalendarNavigation-monthButton--next'
    );
    await RPA.WebBrowser.mouseClick(NextMonthButton);
  }
  if (Number(Month) == Number(FilePath_Month)) {
    await RPA.Logger.info('月が一致しました');
  }
  // ファイル名の日付の部分を取得
  const FilePath_Day = await FilePathData[0].slice(7, -4);
  await RPA.Logger.info(`今日は ${FilePath_Day} 日です`);
  // 1日の位置を検索
  for (let i = 0; i <= 6; i++) {
    const SearchDay: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('Calendar-dayText')[${Number(i)}]`
    );
    const SearchDayText = await SearchDay.getText();
    if (SearchDayText != '1') {
      await RPA.Logger.info(`${SearchDayText} 日`);
      await RPA.Logger.info('1日ではないためスルーします');
      ThroughDay.push(SearchDayText);
    } else if (SearchDayText == '1') {
      await RPA.Logger.info(`${SearchDayText} 日です`);
      await RPA.Logger.info('1列目から判定を開始します');
      break;
    }
  }
  // 1列目の判定
  await FirstJudge(FilePath_Day);
}

const JudgeFlag = ['false'];
async function FirstJudge(FilePath_Day) {
  // 1日から判定開始
  for (let i = ThroughDay.length; i <= 6; i++) {
    const SearchDay2: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('Calendar-dayText')[${Number(i)}]`
    );
    const SearchDayText2 = await SearchDay2.getText();
    if (Number(SearchDayText2) != Number(FilePath_Day)) {
      await RPA.Logger.info(`${SearchDayText2} 日 → 日付が不一致です`);
    } else if (Number(SearchDayText2) == Number(FilePath_Day)) {
      JudgeFlag[0] = 'true';
      await RPA.Logger.info(`${SearchDayText2} 日 → 日付が一致しました`);
      // 一致した日付をクリック
      await RPA.WebBrowser.mouseClick(SearchDay2);
      await RPA.Logger.info('判定処理を終了します');
      break;
    }
  }
  if (JudgeFlag[0] == 'false') {
    await RPA.Logger.info('2列目から判定を再開します');
    // 2列目から判定再開
    await SecondJudge(FilePath_Day);
  }
}

async function SecondJudge(FilePath_Day) {
  for (let i = 7; i <= 41; i++) {
    const SearchDay3: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName('Calendar-dayText')[${Number(i)}]`
    );
    const SearchDayText3 = await SearchDay3.getText();
    if (Number(SearchDayText3) == Number(FilePath_Day)) {
      await RPA.Logger.info(`${SearchDayText3} 日 → 日付が一致しました`);
      await RPA.WebBrowser.mouseClick(SearchDay3);
      await RPA.Logger.info('判定処理を終了します');
      break;
    }
  }
}

async function BrowserReload() {
  const Loading: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('src-components-ErrorMessage-styles-module--loadingErrorMessage')[0].children[0]`
  );
  const LoadingText = await Loading.getText();
  if (LoadingText == 'このコンテンツの読み込み中に問題が発生しました。') {
    await RPA.Logger.info('読み込みエラーが発生したため、ブラウザを更新します');
    await RPA.WebBrowser.refresh();
    await RPA.sleep(5000);
  }
}
