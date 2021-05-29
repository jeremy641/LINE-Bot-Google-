//程式碼開始
//程式碼由　Boris　@　http://www.youtube.com/borispcp　設計，歡迎使用但請保留此聲明。
var　CHANNEL_ACCESS_TOKEN　=　"";
var　formId　=　"";　　　　//表單　ID
var　spreadSheetId　=　"";　　　　//試算表　ID
var　sheetName　=　"";　　　　//工作表名稱
var　destinationFolderID　=　"";　//存放檔案的資料夾ID
var　form　=　FormApp.openById(formId);
var　spreadSheet　=　SpreadsheetApp.openById(spreadSheetId);
var　sheet　=　spreadSheet.getSheetByName(sheetName);
var　lastRow　=　sheet.getLastRow();
var　lastColumn　=　sheet.getLastColumn();
var　sheetData　=　sheet.getSheetValues(1,　1,　lastRow,　lastColumn);
var　confirmationMessageDefault　=　"我們已經收到您回覆的表單。";
var　confirmationMessage　=　form.getConfirmationMessage();
if　(confirmationMessage　==　"")　{confirmationMessage　=　confirmationMessageDefault;}
var　formItems　=　form.getItems();
function　doPost(e)　{
　　var　userData　=　JSON.parse(e.postData.contents);
　　console.log(userData);
　　//　取出　replayToken　和發送的訊息文字
　　var　replyMessage　=　[];
　　var　replyToken　=　userData.events[0].replyToken;
　　var　clientID　=　userData.events[0].source.userId;
　　var　clientMessage;
　　var　nowTime　=　new　Date();
　　var　prefixFileName　=　nowTime.getFullYear()　+　append1Zero(nowTime.getMonth()　+　1)　+　append1Zero(nowTime.getDate())　+　append1Zero(nowTime.getHours())　+　append1Zero(nowTime.getMinutes())　+　append1Zero(nowTime.getSeconds())　+　append2Zero(nowTime.getMilliseconds())　+　"-";
　　var　fileLocation　=　"";
　　switch(userData.events[0].type)　{
　　　　case　"message":
　　　　　　switch(userData.events[0].message.type)　{
　　　　　　　　case　"text":
　　　　　　　　　　clientMessage　=　userData.events[0].message.text;
　　　　　　　　　　break;
　　　　　　　　case　"sticker":
　　　　　　　　case　"location":
　　　　　　　　　　return;
　　　　　　　　　　break;
　　　　　　　　default:
　　　　　　　　　　var　GoogleDrive　=　DriveApp;
　　　　　　　　　　var　saveFolder　=　GoogleDrive.getFolderById(destinationFolderID);
　　　　　　　　　　var　messageId　=　userData.events[0].message.id;
　　　　　　　　　　var　files　=　GoogleDrive.createFile(getFileData(CHANNEL_ACCESS_TOKEN,　messageId).getBlob());
　　　　　　　　　　var　fileExtension　=　files.getName().split(".");
　　　　　　　　　　if　(userData.events[0].message.type　!=　"file")　{
　　　　　　　　　　　　var　GoogleDriveFileName　=　messageId　+　"."　+　fileExtension[fileExtension.length　-　1];
　　　　　　　　　　}
　　　　　　　　　　else　{
　　　　　　　　　　　　var　GoogleDriveFileName　=　userData.events[0].message.fileName;
　　　　　　　　　　}
　　　　　　　　　　var　destinationFile　=　files.makeCopy(prefixFileName　+　clientID　+　"-"　+　GoogleDriveFileName,　saveFolder);
　　　　　　　　　　fileLocation　=　"https://drive.google.com/open?id"　+　destinationFile.getId();
　　　　　　　　　　clientMessage　=　fileLocation;
　　　　　　　　　　GoogleDrive.removeFile(files);
　　　　　　}
　　　　　　break;
　　　　case　"postback":
　　　　　　switch(userData.events[0].postback.data){
　　　　　　　　case　"DateMessage":
　　　　　　　　　　clientMessage　=　userData.events[0].postback.params.date;
　　　　　　　　　　replyMessage.push({type:"text",　text:clientMessage});
　　　　　　　　　　break;
　　　　　　　　case　"TimeMessage":
　　　　　　　　　　clientMessage　=　userData.events[0].postback.params.time;
　　　　　　　　　　replyMessage.push({type:"text",　text:clientMessage});
　　　　　　　　　　break;
　　　　　　　　case　"DateTimeMessage":
　　　　　　　　　　clientMessage　=　userData.events[0].postback.params.datetime;
　　　　　　　　　　replyMessage.push({type:"text",　text:clientMessage});
　　　　　　　　　　break;
　　　　　　　　case　"ignoreQuestion":
　　　　　　　　　　clientMessage　=　"NULL";
　　　　　　　　　　replyMessage.push({type:"text",　text:"此題已略過"});
　　　　　　　　　　break;
　　　　　　　　case　"otherOption":
　　　　　　　　　　replyMessage　=　replyMessage.concat(otherOptionMessage());
　　　　　　　　　　sendReplyMessage(CHANNEL_ACCESS_TOKEN,　replyToken,　replyMessage);
　　　　　　　　　　return;
　　　　　　　　　　break;
　　　　　　}
　　　　　　break;
　　　　case　"follow":
　　　　　　clientMessage　=　"follow";
　　　　　　break;
　　　　default:
　　　　　　return;
　　}
　　var　replyData　=　getUserAnswer(clientID,　clientMessage);
　　switch　(replyData[1])　{
　　　　case　-1:
　　　　　　sheet.getRange(replyData[0],　1).setValue(Date());
　　　　　　replyMessage　=　replyMessage.concat(finishTheQuestionnare(replyData[2]));
　　　　　　sendReplyMessage(CHANNEL_ACCESS_TOKEN,　replyToken,　replyMessage);
　　　　　　return;
　　　　　　break;
　　　　case　1:
　　　　　　replyMessage　=　replyMessage.concat(getFormTitle());
　　　　　　break;
　　}
　　replyMessage　=　replyMessage.concat(getQuestion(replyData[1]));
　　sendReplyMessage(CHANNEL_ACCESS_TOKEN,　replyToken,　replyMessage);
}
//判斷使用者回答到第幾題
function　getUserAnswer(clientID,　clientMessage)　{
　　var　returnData　=　[];
　　for　(var　i　=　lastRow　-　1;　i　>=　0;　i--)　{
　　　　if　(sheetData[i][0]　==　""　&&　sheetData[i][lastColumn　-　1]　==　clientID)　{
　　　　　　for　(var　j　=　1;　j　<=　lastColumn　-1;　j++)　{
　　　　　　　　if　(sheetData[i][j]　===　"")　{break;}
　　　　　　}
　　　　　　sheet.getRange(i　+　1,　j　+　1).setValue(clientMessage);
　　　　　　//如果使用者已經回答了最後一題，就把完成時間填上。不然就送出下一題給使用者
　　　　　　if　(j　+　2　==　lastColumn)　{
　　　　　　　　returnData　=　[i　+　1,　-1,　j];
　　　　　　}
　　　　　　else　{
　　　　　　　　returnData　=　[i　+　1,　j　+　1];
　　　　　　}
　　　　　　return　returnData;
　　　　　　break;
　　　　}
　　}
　　//如果使用者還沒有回答過任何資料，就新增加一列在最後，把使用者ID輸入並開始送出題目
　　sheet.insertRowAfter(lastRow);
　　sheet.getRange(lastRow　+　1,　lastColumn).setValue(clientID);
　　returnData　=　[lastRow　+　1,　1];
　　return　returnData;
}
//取得表單名稱及說明
function　getFormTitle()　{
　　var　formTitleDescription　=　[];
　　var　flexMessage　=　[];
　　flexMessage.push({
　　　　"type":　"bubble",
　　　　"body":　{
　　　　　　"type":　"box",
　　　　　　"layout":　"vertical",
　　　　　　"contents":　[
　　　　　　　　{
　　　　　　　　　　"type":　"text",
　　　　　　　　　　"text":　form.getTitle(),
　　　　　　　　　　"weight":　"bold",
　　　　　　　　　　"size":　"xxl",
　　　　　　　　　　"wrap":　true
　　　　　　　　},
　　　　　　　　{
　　　　　　　　　　"type":　"text",
　　　　　　　　　　"text":　form.getDescription()　+　"　",
　　　　　　　　　　"size":　"xs",
　　　　　　　　　　"color":　"#aaaaaa",
　　　　　　　　　　"wrap":　true,
　　　　　　　　　　"margin":　"md"
　　　　　　　　}
　　　　　　]
　　　　}
　　});
　　formTitleDescription.push({
　　　　"type":　"flex",
　　　　"altText":　form.getTitle(),
　　　　"contents":　{
　　　　　　"type":　"carousel",
　　　　　　"contents":　flexMessage
　　　　}
　　});
　　return　formTitleDescription;
}
//取得要送出的題目
function　getQuestion(questionNo)　{
　　var　replyMessage　=　[];
//如果題目前面有段落，重新計算取得題目真正的　index
　　var　realQuestionNo　=　questionNo;
　　for　(var　i　=　0;　i　<　realQuestionNo;　i++)　{
　　　　if　(formItems[i].getType()　==　"SECTION_HEADER")　{
　　　　　　realQuestionNo++;
　　　　}
　　}
//把題目前面的段落標題及說明取出來　　
　　if　(realQuestionNo　>　questionNo)　{
　　　　var　checkSectionHeaderNo　=　realQuestionNo　-　3;
　　　　var　sectionHeaderTitle　=　[];
　　　　for　(var　i　=　realQuestionNo　-　2;　i　>　checkSectionHeaderNo;　i--)　{
　　　　　　if　(formItems[i]　&&　formItems[i].getType()　==　FormApp.ItemType.SECTION_HEADER)　{
　　　　　　　　checkSectionHeaderNo--;
　　　　　　　　sectionHeaderTitle.push({
　　　　　　　　　　"type":　"flex",
　　　　　　　　　　"altText":　formItems[i].asSectionHeaderItem().getTitle(),
　　　　　　　　　　"contents":　{
　　　　　　　　　　　　"type":　"carousel",
　　　　　　　　　　　　"contents":　[{
　　　　　　　　　　　　　　"type":　"bubble",
　　　　　　　　　　　　　　"body":　{
　　　　　　　　　　　　　　　　"type":　"box",
　　　　　　　　　　　　　　　　"layout":　"vertical",
　　　　　　　　　　　　　　　　"contents":　[
　　　　　　　　　　　　　　　　　　{
　　　　　　　　　　　　　　　　　　　　"type":　"text",
　　　　　　　　　　　　　　　　　　　　"text":　formItems[i].asSectionHeaderItem().getTitle(),
　　　　　　　　　　　　　　　　　　　　"weight":　"bold",
　　　　　　　　　　　　　　　　　　　　"size":　"xl",
　　　　　　　　　　　　　　　　　　　　"wrap":　true
　　　　　　　　　　　　　　　　　　}
　　　　　　　　　　　　　　　　]
　　　　　　　　　　　　　　}
　　　　　　　　　　　　}]
　　　　　　　　　　}
　　　　　　　　});
　　　　　　　　if　(formItems[i].asSectionHeaderItem().getHelpText()　!=　"")　{
　　　　　　　　　　sectionHeaderTitle[sectionHeaderTitle.length　-　1].contents.contents[0].body.contents　=　sectionHeaderTitle[sectionHeaderTitle.length　-　1].contents.contents[0].body.contents.concat(
　　　　　　　　　　　　{
　　　　　　　　　　　　　　"type":　"text",
　　　　　　　　　　　　　　"text":　formItems[i].asSectionHeaderItem().getHelpText(),
　　　　　　　　　　　　　　"size":　"xs",
　　　　　　　　　　　　　　"color":　"#aaaaaa",
　　　　　　　　　　　　　　"wrap":　true,
　　　　　　　　　　　　　　"margin":　"md"
　　　　　　　　　　　　}
　　　　　　　　　　);
　　　　　　　　}
　　　　　　}
　　　　}
　　　　if　(sectionHeaderTitle.length　!=　0)　{replyMessage　=　replyMessage.concat(sectionHeaderTitle.reverse());}　　
　　}
//取得題型
　　var　itemObj　=　formItems[realQuestionNo　-　1];
　　var　itemContent;
　　var　itemFlex　=　[];
　　var　optionsContent　=　[];
　　var　scaleFlexContent　=　[];
　　switch　(itemObj.getType())　{
　　　　case　FormApp.ItemType.MULTIPLE_CHOICE:　　　　//單選題
　　　　case　FormApp.ItemType.LIST:　　　　//下拉式選單
　　　　　　if　(itemObj.getType()　==　FormApp.ItemType.MULTIPLE_CHOICE)　{itemContent　=　itemObj.asMultipleChoiceItem();}
　　　　　　else　{itemContent　=　itemObj.asListItem();}
　　　　　　var　choiceItemOptions　=　itemContent.getChoices();
　　　　　　var　buttonColor;
　　　　　　for　(var　i　=　0;　i　<　choiceItemOptions.length;　i++)　{
　　　　　　　　if　(i　%　2　===　0)　{buttonColor　=　"#9CA3DB";}
　　　　　　　　else　{buttonColor　=　"#677DB7";}
　　　　　　　　optionsContent.push(
　　　　　　　　　　{
　　　　　　　　　　　　"type":　"button",
　　　　　　　　　　　　"action":　{
　　　　　　　　　　　　　　"type":　"message",
　　　　　　　　　　　　　　"label":　choiceItemOptions[i].getValue(),
　　　　　　　　　　　　　　"text":　choiceItemOptions[i].getValue()
　　　　　　　　　　　　},
　　　　　　　　　　　　"style":　"primary",
　　　　　　　　　　　　"color":　buttonColor,
　　　　　　　　　　　　"margin":　"sm"
　　　　　　　　　　}
　　　　　　　　);
　　　　　　}
　　　　　　break;
　　　　case　FormApp.ItemType.SCALE:　　　　　　　　　　　　　　　//線性刻度
　　　　　　itemContent　=　itemObj.asScaleItem();
　　　　　　var　uBound　=　itemContent.getUpperBound();
　　　　　　var　lBound　=　itemContent.getLowerBound();
　　　　　　var　lLabel　=　itemContent.getLeftLabel();
　　　　　　var　rLabel　=　itemContent.getRightLabel();
　　　　　　if　(lLabel　!=　"")　{
　　　　　　　　scaleFlexContent.push(
　　　　　　　　　　{
　　　　　　　　　　　　"type":　"bubble",
　　　　　　　　　　　　"size":　"nano",
　　　　　　　　　　　　"body":　{
　　　　　　　　　　　　　　"type":　"box",
　　　　　　　　　　　　　　"layout":　"horizontal",
　　　　　　　　　　　　　　"contents":　[
　　　　　　　　　　　　　　　　{
　　　　　　　　　　　　　　　　　　"type":　"text",
　　　　　　　　　　　　　　　　　　"text":　lLabel,
　　　　　　　　　　　　　　　　　　"align":　"center",
　　　　　　　　　　　　　　　　　　"gravity":　"center",
　　　　　　　　　　　　　　　　　　"wrap":　true,
　　　　　　　　　　　　　　　　　　"size":　"xl"
　　　　　　　　　　　　　　　　}
　　　　　　　　　　　　　　]
　　　　　　　　　　　　},
　　　　　　　　　　　　"styles":　{
　　　　　　　　　　　　　　"footer":　{
　　　　　　　　　　　　　　　　"separator":　false
　　　　　　　　　　　　　　}
　　　　　　　　　　　　}
　　　　　　　　　　}
　　　　　　　　)
　　　　　　}
　　　　　　for　(var　i　=　lBound;　i　<=　uBound;　i++)　{
　　　　　　　　scaleFlexContent.push(
　　　　　　　　　　{
　　　　　　　　　　　　"type":　"bubble",
　　　　　　　　　　　　"size":　"nano",
　　　　　　　　　　　　"body":　{
　　　　　　　　　　　　　　"type":　"box",
　　　　　　　　　　　　　　"layout":　"horizontal",
　　　　　　　　　　　　　　"contents":　[
　　　　　　　　　　　　　　　　{
　　　　　　　　　　　　　　　　　　"type":　"button",
　　　　　　　　　　　　　　　　　　"action":　{
　　　　　　　　　　　　　　　　　　　　"type":　"message",
　　　　　　　　　　　　　　　　　　　　"label":　i.toString(),
　　　　　　　　　　　　　　　　　　　　"text":　i.toString()
　　　　　　　　　　　　　　　　　　},
　　　　　　　　　　　　　　　　　　"style":　"primary",
　　　　　　　　　　　　　　　　　　"color":　"#9CA3DB",
　　　　　　　　　　　　　　　　　　"gravity":　"center"
　　　　　　　　　　　　　　　　}
　　　　　　　　　　　　　　]
　　　　　　　　　　　　},
　　　　　　　　　　　　"styles":　{
　　　　　　　　　　　　　　"footer":　{
　　　　　　　　　　　　　　　"separator":　false
　　　　　　　　　　　　　　}
　　　　　　　　　　　　}
　　　　　　　　　　}
　　　　　　　　)
　　　　　　}
　　　　　　if　(rLabel　!=　"")　{
　　　　　　　　scaleFlexContent.push(
　　　　　　　　　　{
　　　　　　　　　　　　"type":　"bubble",
　　　　　　　　　　　　"size":　"nano",
　　　　　　　　　　　　"body":　{
　　　　　　　　　　　　　　"type":　"box",
　　　　　　　　　　　　　　"layout":　"horizontal",
　　　　　　　　　　　　　　"contents":　[
　　　　　　　　　　　　　　　　{
　　　　　　　　　　　　　　　　　　"type":　"text",
　　　　　　　　　　　　　　　　　　"text":　rLabel,
　　　　　　　　　　　　　　　　　　"align":　"center",
　　　　　　　　　　　　　　　　　　"gravity":　"center",
　　　　　　　　　　　　　　　　　　"wrap":　true,
　　　　　　　　　　　　　　　　　　"size":　"xl"
　　　　　　　　　　　　　　　　}
　　　　　　　　　　　　　　]
　　　　　　　　　　　　},
　　　　　　　　　　　　"styles":　{
　　　　　　　　　　　　　　"footer":　{
　　　　　　　　　　　　　　　　"separator":　false
　　　　　　　　　　　　　　}
　　　　　　　　　　　　}
　　　　　　　　　　}
　　　　　　　　)
　　　　　　}
　　　　　　break;
　　　　case　FormApp.ItemType.TEXT:　　　　　　　　　　　　　　　//簡答題
　　　　　　itemContent　=　itemObj.asTextItem();
　　　　　　break;
　　　　case　FormApp.ItemType.PARAGRAPH_TEXT:　　　　　//段落問題
　　　　　　itemContent　=　itemObj.asParagraphTextItem();
　　　　　　break;
　　　　case　FormApp.ItemType.DATE:　　　　　　　　　　　　　　　//日期
　　　　　　itemContent　=　itemObj.asDateItem();
　　　　　　optionsContent.push(
　　　　　　　　{
　　　　　　　　　　"type":　"button",
　　　　　　　　　　"action":　{
　　　　　　　　　　　　"type":　"datetimepicker",
　　　　　　　　　　　　"label":　"點選輸入日期",
　　　　　　　　　　　　"data":　"DateMessage",
　　　　　　　　　　　　"mode":　"date"
　　　　　　　　　　},
　　　　　　　　　　"style":　"primary",
　　　　　　　　　　"color":　"#454B66",
　　　　　　　　　　"margin":　"sm"
　　　　　　　　}
　　　　　　);
　　　　　　break;
　　　　case　FormApp.ItemType.TIME:　　　　　　　　　　　　　　　//時間
　　　　　　itemContent　=　itemObj.asTimeItem();
　　　　　　optionsContent.push(
　　　　　　　　{
　　　　　　　　　　"type":　"button",
　　　　　　　　　　"action":　{
　　　　　　　　　　　　"type":　"datetimepicker",
　　　　　　　　　　　　"label":　"點選輸入時間",
　　　　　　　　　　　　"data":　"TimeMessage",
　　　　　　　　　　　　"mode":　"time"
　　　　　　　　　　},
　　　　　　　　　　"style":　"primary",
　　　　　　　　　　"color":　"#454B66",
　　　　　　　　　　"margin":　"sm"
　　　　　　　　}
　　　　　　);
　　　　　　break;
　　　　case　FormApp.ItemType.DATETIME:　　　　　　　　　　//日期及時間
　　　　　　itemContent　=　itemObj.asDateTimeItem();
　　　　　　optionsContent.push(
　　　　　　　　{
　　　　　　　　　　"type":　"button",
　　　　　　　　　　"action":　{
　　　　　　　　　　　　"type":　"datetimepicker",
　　　　　　　　　　　　"label":　"點選輸入日期及時間",
　　　　　　　　　　　　"data":　"DateTimeMessage",
　　　　　　　　　　　　"mode":　"datetime"
　　　　　　　　　　},
　　　　　　　　　　"style":　"primary",
　　　　　　　　　　"color":　"#454B66",
　　　　　　　　　　"margin":　"sm"
　　　　　　　　}
　　　　　　);
　　　　　　break;
　　　　case　FormApp.ItemType.FILE_UPLOAD:　　　　　　　　//上傳檔案
　　　　　　itemContent　=　itemObj;
　　　　　　break;
　　}
　　if　(itemObj.getType()　==　FormApp.ItemType.MULTIPLE_CHOICE　&&　itemContent.hasOtherOption())　{
　　　　optionsContent.push(
　　　　　　{
　　　　　　　　"type":　"button",
　　　　　　　　"action":　{
　　　　　　　　　　"type":　"postback",
　　　　　　　　　　"label":　"其他",
　　　　　　　　　　"data":　"otherOption"
　　　　　　　　},
　　　　　　　　"style":　"primary",
　　　　　　　　"color":　"#454B66",
　　　　　　　　"margin":　"sm"
　　　　　　}
　　　　);
　　}
　　if　(itemObj.getType()　!=　"FILE_UPLOAD")　{
　　　　if　(!itemContent.isRequired())　{
　　　　　　optionsContent.push(
　　　　　　　　{
　　　　　　　　　　"type":　"button",
　　　　　　　　　　"action":　{
　　　　　　　　　　　　"type":　"postback",
　　　　　　　　　　　　"label":　"略過此題",
　　　　　　　　　　　　"data":　"ignoreQuestion"
　　　　　　　　　　},
　　　　　　　　　　"margin":　"sm"
　　　　　　　　}
　　　　　　);
　　　　}
　　}
　　itemFlex.push({
　　　　"type":　"flex",
　　　　"altText":　itemObj.getTitle(),
　　　　"contents":　{
　　　　　　"type":　"carousel",
　　　　　　"contents":　[{
　　　　　　　　"type":　"bubble",
　　　　　　　　"body":　{
　　　　　　　　　　"type":　"box",
　　　　　　　　　　"layout":　"vertical",
　　　　　　　　　　"contents":　[
　　　　　　　　　　　　{
　　　　　　　　　　　　　　"type":　"text",
　　　　　　　　　　　　　　"text":　itemObj.getTitle(),
　　　　　　　　　　　　　　"weight":　"bold",
　　　　　　　　　　　　　　"size":　"xl",
　　　　　　　　　　　　　　"wrap":　true
　　　　　　　　　　　　}
　　　　　　　　　　]
　　　　　　　　}
　　　　　　}]
　　　　}
　　});
　　if　(itemObj.getHelpText()　!=　"")　{
　　　　itemFlex[0].contents.contents[0].body.contents　=　itemFlex[0].contents.contents[0].body.contents.concat({
　　　　　　　　"type":　"text",
　　　　　　　　"text":　itemObj.getHelpText(),
　　　　　　　　"size":　"xs",
　　　　　　　　"color":　"#aaaaaa",
　　　　　　　　"wrap":　true,
　　　　　　　　"margin":　"md"
　　　　　　}
　　　　);
　　}
　　if　(optionsContent.length　!=　0)　{
　　　　itemFlex[0].contents.contents[0].body.contents　=　itemFlex[0].contents.contents[0].body.contents.concat({
　　　　　　　　"type":　"separator",
　　　　　　　　"margin":　"xxl"
　　　　　　},
　　　　　　{
　　　　　　　　"type":　"box",
　　　　　　　　"layout":　"vertical",
　　　　　　　　"margin":　"md",
　　　　　　　　"contents":　optionsContent
　　　　　　}
　　　　);
　　}
　　if　(scaleFlexContent.length　!=　0)　{
　　　　itemFlex.push({
　　　　　　"type":　"flex",
　　　　　　"altText":　itemObj.getTitle(),
　　　　　　"contents":　{
　　　　　　　　"type":　"carousel",
　　　　　　　　"contents":scaleFlexContent
　　　　　　}
　　　　});
　　}
　　replyMessage　=　replyMessage.concat((itemFlex));
　　return　replyMessage;
}
function　otherOptionMessage()　{
　　var　returnData　=　[];
　　returnData.push({
　　　　"type":　"flex",
　　　　"altText":　"請輸入「其他」的內容",
　　　　"contents":　{
　　　　　　"type":　"carousel",
　　　　　　"contents":　[{
　　　　　　　　"type":　"bubble",
　　　　　　　　"body":　{
　　　　　　　　　　"type":　"box",
　　　　　　　　　　"layout":　"vertical",
　　　　　　　　　　"contents":　[
　　　　　　　　　　　　{
　　　　　　　　　　　　　　"type":　"text",
　　　　　　　　　　　　　　"text":　"您選擇了「其他」，麻煩請輸入您的答案後送出即可",
　　　　　　　　　　　　　　"weight":　"bold",
　　　　　　　　　　　　　　"size":　"xl",
　　　　　　　　　　　　　　"wrap":　true
　　　　　　　　　　　　}
　　　　　　　　　　]
　　　　　　　　}
　　　　　　}]
　　　　}
　　});
　　return　returnData;
}
//回送　Line　Bot　訊息給使用者
function　sendReplyMessage(CHANNEL_ACCESS_TOKEN,　replyToken,　replyMessage)　{
　　var　url　=　"https://api.line.me/v2/bot/message/reply";
　　UrlFetchApp.fetch(url,　{
　　　　"headers":　{
　　　　　　"Content-Type":　"application/json;　charset=UTF-8",
　　　　　　"Authorization":　"Bearer　"　+　CHANNEL_ACCESS_TOKEN,
　　　　},
　　　　"method":　"post",
　　　　"payload":　JSON.stringify({
　　　　　　"replyToken":　replyToken,
　　　　　　"messages":　replyMessage,
　　　　}),
　　});
}
//取得檔案的　Binary　資料
function　getFileData(CHANNEL_ACCESS_TOKEN,　fileID){
　　var　url　=　"https://api.line.me/v2/bot/message/"　+　fileID　+　"/content";
　　return　UrlFetchApp.fetch(url,　{
　　　　'headers':　{
　　　　　　'Authorization':　'Bearer　'　+　CHANNEL_ACCESS_TOKEN,
　　　　},
　　　　'method':　'get',
　　});
}
//取得最後一個問題之後的東西以及確認訊息
function　finishTheQuestionnare(lastNum)　{
　　var　replyMessage　=　[];
//如果題目前面有段落，重新計算取得題目真正的　index
　　var　realQuestionNo　=　lastNum;
　　for　(var　i　=　0;　i　<　realQuestionNo;　i++)　{
　　　　if　(formItems[i].getType()　==　"SECTION_HEADER")　{
　　　　　　realQuestionNo++;
　　　　}
　　}
　　if　(realQuestionNo　<　formItems.length)　{
　　　　for　(var　i　=　realQuestionNo;　i　<　formItems.length;　i++)　{
　　　　　　replyMessage.push({
　　　　　　　　"type":　"flex",
　　　　　　　　"altText":　formItems[i].getTitle(),
　　　　　　　　"contents":　{
　　　　　　　　　　"type":　"carousel",
　　　　　　　　　　"contents":　[{
　　　　　　　　　　　　"type":　"bubble",
　　　　　　　　　　　　"body":　{
　　　　　　　　　　　　　　"type":　"box",
　　　　　　　　　　　　　　"layout":　"vertical",
　　　　　　　　　　　　　　"contents":　[
　　　　　　　　　　　　　　　　{
　　　　　　　　　　　　　　　　　　"type":　"text",
　　　　　　　　　　　　　　　　　　"text":　formItems[i].getTitle(),
　　　　　　　　　　　　　　　　　　"weight":　"bold",
　　　　　　　　　　　　　　　　　　"size":　"xl",
　　　　　　　　　　　　　　　　　　"wrap":　true
　　　　　　　　　　　　　　　　}
　　　　　　　　　　　　　　]
　　　　　　　　　　　　}
　　　　　　　　　　}]
　　　　　　　　}
　　　　　　});
　　　　　　if　(formItems[i].getHelpText()　!=　"")　{
　　　　　　　　replyMessage[replyMessage.length　-　1].contents.contents[0].body.contents　=　replyMessage[replyMessage.length　-　1].contents.contents[0].body.contents.concat(
　　　　　　　　　　{
　　　　　　　　　　　　"type":　"text",
　　　　　　　　　　　　"text":　formItems[i].getHelpText(),
　　　　　　　　　　　　"size":　"xs",
　　　　　　　　　　　　"color":　"#aaaaaa",
　　　　　　　　　　　　"wrap":　true,
　　　　　　　　　　　　"margin":　"md"
　　　　　　　　　　}
　　　　　　　　);
　　　　　　}
　　　　}
　　}
　　replyMessage.push({
　　　　"type":　"flex",
　　　　"altText":　confirmationMessage,
　　　　"contents":　{
　　　　　　"type":　"carousel",
　　　　　　"contents":　[{
　　　　　　　　"type":　"bubble",
　　　　　　　　"body":　{
　　　　　　　　　　"type":　"box",
　　　　　　　　　　"layout":　"vertical",
　　　　　　　　　　"contents":　[
　　　　　　　　　　　　{
　　　　　　　　　　　　　　"type":　"text",
　　　　　　　　　　　　　　"text":　confirmationMessage,
　　　　　　　　　　　　　　"weight":　"bold",
　　　　　　　　　　　　　　"size":　"xl",
　　　　　　　　　　　　　　"wrap":　true
　　　　　　　　　　　　}
　　　　　　　　　　]
　　　　　　　　}
　　　　　　}]
　　　　}
　　});
　　return　replyMessage;
}
//前綴補「0」函式
function　append1Zero(obj)　{
　　if　(obj　<　10)　{return　"0"　+　obj.toString();}
　　else　{return　obj.toString();}
}
function　append2Zero(obj)　{
　　if　(obj　<　10)　{return　"00"　+　obj.toString();}
　　else　if　(obj　<　100　&&　obj　>=　10)　{return　"0"　+　obj.toString();}
　　else　{return　obj.toString();}
}
//程式碼結束