import * as LineAppLib from "../../line_app/src/line_app";
import { WebhookEvent, TemplateMessage } from "@line/bot-sdk";
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

//LINEからのイベントがdoPostにとんでくる
function doPost(e) {
  const spreadSheetId =
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const spreadSheet = SpreadsheetApp.openById(spreadSheetId);
  const debugsheet = spreadSheet.getSheetByName("debug");
  const token =
    PropertiesService.getScriptProperties().getProperty("LINE_ACCESS_TOKEN");
  const lineApp = new LineAppLib.LineApp(token);
  //とんできた情報を扱いやすいように変換している
  const json: string = e.postData.contents;
  const events: WebhookEvent[] = JSON.parse(json).events;

  //とんできたイベントの種類を確認する
  events.forEach(function (event: WebhookEvent) {
    //debugsheet.appendRow([userId]);
    if (event.type == "postback") {
      lineApp.sendSimpleReplyMessage(
        event.replyToken,
        "申し訳ありませんが、Postbackメッセージには現在対応しておりません。",
      );
    } else if (event.type == "message") {
      if (event.message.type == "image") {
        // message/image の場合グループチャットのみ画像保存を行う.
        if (event.source.type == "group") {
          const groupFolderName = JSON.parse(
            getGroupSummaryResponse(event.source.groupId).getContentText(),
          ).groupName;
          const imageBlob = getImage(event.message.id);
          const groupFolder = getSubFolderOrCreateIfNotExist(groupFolderName);
          saveImageBlobAsPng(imageBlob, groupFolder);
        } else {
          lineApp.sendSimpleReplyMessage(
            event.replyToken,
            "このBOTはグループチャット専用となっております。グループチャットに友達登録してご利用ください。",
          );
        }
      } else if (event.message.type == "file") {
        if (event.source.type == "group") {
          const groupFolderName = JSON.parse(
            getGroupSummaryResponse(event.source.groupId).getContentText(),
          ).groupName;
          const imageBlob = getFile(event.message.id);
          const groupFolder = getSubFolderOrCreateIfNotExist(groupFolderName);
          saveImageBlobAsPng(imageBlob, groupFolder);
        } else {
          lineApp.sendSimpleReplyMessage(
            event.replyToken,
            "このBOTはグループチャット専用となっております。グループチャットに友達登録してご利用ください。",
          );
        }
      } else if (event.message.type == "text") {
        // message/text の場合個人トークのみ画像検索を行う.
        if (event.source.type == "group") {
          // 何もしない
        } else if (event.source.type == "user") {
          var testArr = [];
          const fileUrl = searchFiles(textToQueryString(event.message.text));
          if (fileUrl.length == 0) {
            lineApp.sendSimpleReplyMessage(
              event.replyToken,
              `「${event.message.text}」にマッチしたスライドが見つからないようです。キーワードを変更してみてください。`,
            );
          } else if (fileUrl.length >= 1 && fileUrl.length <= 5) {
            fileUrl.forEach(function (url) {
              testArr.push({
                type: "image",
                originalContentUrl: url,
                previewImageUrl: url,
              });
            });
          } else {
            testArr.push({
              type: "text",
              text: `「${event.message.text}」では候補が多すぎるため、見つかったスライドのうちランダムに４つのみお送りします。複数のキーワードをスペースで区切って使ってみてください。(例: 「COVID19 皮疹」)`,
            });
            for (var i = 0; i < 4; i++) {
              const url = fileUrl[i];
              testArr.push({
                type: "image",
                originalContentUrl: url,
                previewImageUrl: url,
              });
            }
          }

          lineApp.sendReplyMessageArray(event.replyToken, testArr);
          var dt = new Date();
          var d1 = Utilities.formatDate(
            dt,
            "Asia/Tokyo",
            "yyyy-MM-dd_hh:mm:dd",
          );
          debugsheet.appendRow([
            d1,
            "user",
            "searchFiles",
            event.source.userId,
            event.message.text,
            JSON.stringify(testArr),
          ]);
        }
      } else {
        // 何もしない
      }
    }
  });
}

/**
 * テキスト情報をDrive.apiのファイル検索に用いるクエリストリングへ変換する.
 * テキストはスペース区切りで、fullTextとTitleについてAND検索を行うものとする.
 * @param input スペース区切りの検索用キーワード
 * @returns Drive.apiのクエリストリング
 */
function textToQueryString(input: string): string {
  try {
    // テスト実行用input定義
    var input: string;
    if (!input) {
      Logger.log("textToQueryString() test start");
      input = "COVID   皮疹";
    } else {
      // 引数で与えられたinputを使用する.
    }
    // inputの空白文字処理
    const inputs = input.replace(/\s\s+/g, " ").split(/\s/); // 空白文字をスペース一つに変換
    // outputの生成
    const outputArr = [];
    inputs.forEach(input => {
      outputArr.push(
        `(fullText contains "${input}" or title contains "${input}")`,
      );
    });
    outputArr.push(`(mimeType contains "image/")`);
    const queryString = outputArr.join(" and ");
    console.log(queryString);
    return queryString;
  } catch (e) {
    console.error(`textToQueryString() ERROR: ${e}`);
  }
}

/**
 * Drive.apiを使って画像ファイル検索を行う
 * @param queryString Drive.api用のクエリストリング
 * @returns 画像共有用URLの配列
 */
function searchFiles(queryString: string): string[] {
  try {
    // テスト実行用input定義
    var queryString: string;
    if (!queryString) {
      Logger.log("searchFiles() test start");
      queryString = 'fullText contains "COVID" and fullText contains "皮疹"';
    } else {
      // 引数で与えられたqueryStringを使用する.
    }
    // 最上位のフォルダを取得する
    const gDriveID =
      PropertiesService.getScriptProperties().getProperty("GDRIVE_ID");
    const srcFolder = DriveApp.getFolderById(gDriveID);
    const srcFolders = srcFolder.getFolders(); //フォルダ内フォルダをゲット
    // ファイルを検索する
    const imageUrlArray: string[] = [];
    while (srcFolders.hasNext()) {
      const nextSrcFolder = srcFolders.next();
      const files = nextSrcFolder.searchFiles(queryString); // クエリストリングによるファイル検索
      while (files.hasNext()) {
        const file = files.next();
        const fileId = file.getId();
        // 画像共有用URLを取得しArrayに格納する.
        imageUrlArray.push(`https://drive.google.com/uc?id=${fileId}`);
      }
    }
    console.log(imageUrlArray);
    return imageUrlArray;
  } catch (e) {
    console.error(`searchFiles() ERROR: ${e}`);
  }
}

/**
 * LINEで送信された画像をGASで利用できるように再取得する.
 * 具体的には、image messageのmessageIdを元にGoogle driveに保存可能なimageBlobとして取得している.
 * @param messageId image message の messageId
 * @returns 画像ファイルのBlob. デフォルトではimage/png形式のtemp.pngという名称.
 */
function getFile(messageId: string): GoogleAppsScript.Base.Blob {
  try {
    // image messageのcontentをFetchして取得する.
    const token =
      PropertiesService.getScriptProperties().getProperty("LINE_ACCESS_TOKEN");
    const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
    const options: URLFetchRequestOptions = {
      method: "get",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + token,
      },
    };
    const res = UrlFetchApp.fetch(url, options);
    // 画像ファイルのBlobを取得して返り値とする. デフォルトではimage/png形式のtemp.pngという名称.
    const blob = res.getBlob().getAs("application/pdf").setName("temp.pdf");
    console.log("imageBlobの取得に成功しました");
    console.log("ContentType:" + blob.getContentType());
    console.log("Name: " + blob.getName());
    return blob;
  } catch (e) {
    console.error(`Error: getFile() ${e}`);
  }
}

/**
 * LINEで送信された画像をGASで利用できるように再取得する.
 * 具体的には、image messageのmessageIdを元にGoogle driveに保存可能なimageBlobとして取得している.
 * @param messageId image message の messageId
 * @returns 画像ファイルのBlob. デフォルトではimage/png形式のtemp.pngという名称.
 */
function getImage(messageId: string): GoogleAppsScript.Base.Blob {
  try {
    // image messageのcontentをFetchして取得する.
    const token =
      PropertiesService.getScriptProperties().getProperty("LINE_ACCESS_TOKEN");
    const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
    const options: URLFetchRequestOptions = {
      method: "get",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + token,
      },
    };
    const res = UrlFetchApp.fetch(url, options);
    // 画像ファイルのBlobを取得して返り値とする. デフォルトではimage/png形式のtemp.pngという名称.
    const imageBlob = res.getBlob().getAs("image/png").setName("temp.png");
    console.log("imageBlobの取得に成功しました");
    console.log("ContentType:" + imageBlob.getContentType());
    console.log("Name: " + imageBlob.getName());
    return imageBlob;
  } catch (e) {
    console.error(`Error: getImage() ${e}`);
  }
}
/**
 * BOTが所属しているグループチャットのグループサマリーを取得する.
 * @param groupId BOTが所属しているグループチャットのグループID
 * @returns {GoogleAppsScript.URL_Fetch.HTTPResponse} グループサマリーのHTTPResponse
 *
 * ```typescript
 * const groupName = JSON.parse(getGroupSummaryResponse(event.source.groupId).getContentText()).groupName;
 * ```
 */
function getGroupSummaryResponse(
  groupId: string,
): GoogleAppsScript.URL_Fetch.HTTPResponse {
  try {
    // BOTが所属しているグループチャットのグループサマリーを取得する.
    const token =
      PropertiesService.getScriptProperties().getProperty("LINE_ACCESS_TOKEN");
    const url = `https://api.line.me/v2/bot/group/${groupId}/summary`;
    const options: URLFetchRequestOptions = {
      method: "get",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + token,
      },
    };
    const res = UrlFetchApp.fetch(url, options);
    console.log("グループサマリー取得成功");
    return res;
  } catch (e) {
    console.error(`Error: getGroupSummaryResponse() ${e}`);
  }
}

/**
 * DriveのFolderを指定して、そこにimageのBlobを保存してファイル名をタイムスタンプからつけます.
 * @param imageBlob imageのBlob
 * @param folder 保存先DriveのFolder
 * @returns {GoogleAppsScript.Drive.File}
 */
function saveImageBlobAsPng(
  imageBlob: GoogleAppsScript.Base.Blob,
  folder: GoogleAppsScript.Drive.Folder,
): GoogleAppsScript.Drive.File {
  try {
    const file = folder.createFile(imageBlob);
    const dt = new Date();
    const d1 = Utilities.formatDate(dt, "Asia/Tokyo", "yyyy-MM-dd_hh:mm:dd");
    file.setName(d1);
    console.log(
      "[INFO] Google Driveの以下のURLに画像が保存されました: " +
        folder.getUrl(),
    );
    return file;
  } catch (e) {
    console.error(`Error: saveImageBlobAsPng() ${e}`);
  }
}

/**
 * ソースファイル直下のサブフォルダを名前で指定する。もしサブフォルダが存在した場合サブフォルダを返す。存在しない場合は新しく作成する.
 * @param subFolderName ソースファイル直下のサブフォルダ名称
 * @returns
 */
function getSubFolderOrCreateIfNotExist(
  subFolderName: string,
): GoogleAppsScript.Drive.Folder {
  try {
    // ソースフォルダへアクセス
    const gDriveID =
      PropertiesService.getScriptProperties().getProperty("GDRIVE_ID");
    const srcFolder = DriveApp.getFolderById(gDriveID);
    // ソースフォルダに格納されているアクティブなフォルダ内フォルダを獲得
    const srcFolders = srcFolder.getFolders();
    var FolderNotExistsFlag = true; // フォルダが存在しないフラグ
    while (srcFolders.hasNext()) {
      const nextSrcFolder = srcFolders.next();
      const strFolder = nextSrcFolder.getName();
      if (strFolder == subFolderName) {
        // 同一フォルダ名があればフラグをFalseに変更してそのフォルダを返す
        FolderNotExistsFlag = false;
        return nextSrcFolder;
      } else {
        // 何もしない
      }
    }
    // 指定したフォルダが存在しない場合は新規フォルダを作成する
    if (FolderNotExistsFlag) {
      const newFolder = srcFolder.createFolder(subFolderName);
      return newFolder;
    } else {
      console.error(
        `getSubFolderOrCreateIfNotExist()でフォルダ${subFolderName}を作成しようとしましたが失敗しました`,
      );
    }
  } catch (e) {
    console.error(
      `getSubFolderOrCreateIfNotExist()でフォルダ${subFolderName}を作成または獲得しようとしましたが失敗しました.${e}`,
    );
  }
}
