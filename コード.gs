/**
* スライドタイトル一括変更
* @Author        :kawagoe
* @Version       :1.0
* @Create        :0.1  新規作成
* @Update        :1.0  メインバージョンリリース
* @Etc           :
* @Reference     :
*/

var SHEET, SLIDE;
var Cfg = {
  slideId : "",   //sample "1qNNgRW6649mjGDPhAeaZARlBMuqtCjTgoHGMQjQlN08"
  stRow : 2,
  stCol : 3,
  idIndex : "A1",
  idx  : {
    execFlag    : 0,
    renameTitle : 1,
    shapeId     : 6
  },
  status : [true,false]
};

function onOpen(e){
  SpreadsheetApp.getActive().toast("Access OK ","INFO",5);
  Logger.log(JSON.stringify(e));
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CustomMenu')
  .addItem('タイトル一覧取得', 'getSlideInfo')
  .addItem('タイトルリネーム', 'renameTitle')
  .addToUi();
}

function initial(){
  try{
    SA = SpreadsheetApp.getActive();
    SHEET = SpreadsheetApp.getActiveSheet();
    Cfg.slideId = SHEET.getRange("A1").getValue();
    SLIDE = SlidesApp.openById(Cfg.slideId);
  }catch(e){
    return false;
  }
  return true;
}

/**
 * read a slide.
 * @param {string} none
 */
function getSlideInfo(){
  var ret = initial();
  if ( ret == false ) {
    SA.toast("読み取りエラー","ERROR",5);
    return;
  }
  
  var inputArr = [['Title','PlaceholderType','TextType','FontSize(null=anyStyle)','ShapesId','PageId','PageUrl','Type']];
  var enRow = SHEET.getLastRow() - Cfg.stRow + 1;
  if ( enRow <= Cfg.stRow ) {
    enRow = SHEET.getLastRow();
  }
  
  var　pType, pageId, objId, elType, pUrl, title, tType, fSize;
  var baseUrl = SLIDE.getUrl();
  var slides = SLIDE.getSlides();
  for (var i = 0; i < slides.length; i++) {
//    Logger.log(url + "#slide=id." + slides[i].getObjectId());
    //    Logger.log(pe.length);
//    var pe = slides[i].getPageElements();
    var pe = slides[i].getShapes();
    pageId = slides[i].getObjectId();
    
    for (var j = 0; j < pe.length; j++) {
      try{
        objId = pe[j].getObjectId();
        pUrl = baseUrl + "#slide=id." + pageId;
        pType = pe[j].getPlaceholderType();
        elType = pe[j].getPageElementType();
        tType = pe[j].getShapeType();
        fSize = pe[j].getText().getTextStyle().getFontSize();
        
//        Logger.log(checkType);
        //CENTERED_TITLE / TITLE / SUBTITLE / BODY / NONE
        switch (pType){
          case pType.CENTERED_TITLE :
          case pType.TITLE :
          case pType.SUBTITLE :
            Logger.log(pType + " : " + fSize);
        
            title = pe[j].getText().asString();
            break;
          default :
            continue;
        };
        
//        Logger.log("[" + pe[j].getObjectId() + "]" + "[" + pe[j].getPageElementType() + "]" + title);
        inputArr.push(
          [
            title,
            pType,
            tType,
            fSize,
            objId,
            pageId, 
            pUrl,
            elType
          ]);
      }catch(e){
        LogSheet("ERROR",e.lineNumber + ":" + e.message);
      } //try
    } //for j
  } //for i
  
  Logger.log(inputArr[0].length - Cfg.stCol + 1);
  LogSheet("INFO","取得処理終了");
  SHEET.getRange(Cfg.stRow, 1, enRow, SHEET.getLastColumn()).clear();
  SpreadsheetApp.flush();
  SHEET.getRange(1, Cfg.stCol, inputArr.length, inputArr[0].length).setValues(inputArr);
  
}


/**
 * rename slide title.
 * @param {string} none
 */
function renameTitle() {
  var ret = initial();
  if ( ret == false ) {
    SA.toast("読み取りエラー","ERROR",5);
    return;
  }
  
  var shapeId, renameTitle
  var cnt = 0;
  var inputFlag = [];
  var enRow = SHEET.getLastRow() - Cfg.stRow + 1;
  if ( enRow <= Cfg.stRow ) {
    enRow = SHEET.getLastRow() + 1;
  }
  var enCol = SHEET.getLastColumn() - Cfg.stCol + 1;
  var sData = SHEET.getRange(Cfg.stRow, 1, enRow, enCol).getValues();
  sData.forEach(function(val){
    if (val[0] == true ) cnt++
  });
  
  var ans = Browser.msgBox("実行確認", "A列にチェックが入っているものを対象にリネームします。\\n\\n本当によろしいですか？\\n対象：" + cnt + "件", Browser.Buttons.OK_CANCEL);
  if ( ans != "ok" ) {
    SA.toast("キャンセルしました。処理を終了しましす。","INFO",5);
    return
  }

  var url = SLIDE.getUrl();
  var slides = SLIDE.getSlides();
  
  cnt = 0
  for ( var i = 0; i < sData.length; i++){
    try{
      inputFlag.push([false]);
      
      if ( sData[i][Cfg.idx.execFlag] != true ) {
        continue;
      }
      
      
      //CENTERED_TITLE / TITLE / SUBTITLE / BODY / NONE
      shapeId = sData[i][Cfg.idx.shapeId];
      renameTitle = sData[i][Cfg.idx.renameTitle];
      
      SLIDE.getPageElementById(shapeId).asShape().getText().setText(renameTitle);
      cnt++;
    }catch(e){
      LogSheet("ERROR",e.lineNumber + ":" + e.message);
    } //try
  } //for
  
  if ( 0 < cnt ){
    SHEET.getRange(Cfg.stRow, 1, inputFlag.length, 1).setValues(inputFlag);
  }else{
    SA.toast("対象なし","INFO",5);
  }
  LogSheet("INFO","リネーム処理終了");
}

