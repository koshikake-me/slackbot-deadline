/*=== TYPES ===*/
// 通知する時間についての文章
const RemindType = {
  oneDayBefore: '明日',
  today: '本日',
}
Object.freeze(RemindType);
// 予定について特別な意味をもつ文章
const MagicType = {
  done: 'DONE',
  tbd: 'TBD',
}
Object.freeze(MagicType);
// カテゴリについての定義
const CategoryType = {
  postReport: 'リポートの提出',
  postMediaReport: 'Mスク用リポートの提出',
  registerTanshu: '単修試験の登録',
  takeTanshu: '単修試験の実施日',
  drawLotForSchooling: 'スクーリング抽選の申込',
  transferOfExpenses: '費用の振込',
  proceduresGraduationTheses: '卒業論文関係の手続',
  applicationForMidYearFraduation: '年度途中卒業の申請',
}
Object.freeze(CategoryType);
// 通知するタイミングについての定義
const NoticeType = {
  takeTanshu: '0',
  startDay: '1',
  endBeforeOneDay: '2',
  endDay: '3',
}
Object.freeze(NoticeType);
// 時間についての文章
const DateType = {
  ONE_DAY: 1000*60*60*24*1,
}
Object.freeze(DateType);

/*=== CONSTANTS ===*/
// Slackのトークン
const SLACK_TOKEN = 'INSERT_YOUR_API_KEI_HERE';
// Botがチャットするチャンネル
const TARGET_CHANNEL = '8-02_各種締め切り予定表';
// 暫定的なのかもしれない
const TARGET_CHANNEL_ID = 'INSERT_YOUR_TARGET_CHANNEL_ID_HERE';
// 関連づけたスプレッドシート
const G_SHEET = SpreadsheetApp.getActiveSpreadsheet();

/*=== GOOGLE SHEETS ===*/
// メッセージの素材が書いてある
const utilSheet = G_SHEET.getSheetByName('util');
// 以下は各カテゴリについてのシート
const postReportSheet
  = G_SHEET.getSheetByName(CategoryType.postReport);
const postMediaReportSheet
  = G_SHEET.getSheetByName(CategoryType.postMediaReport);
const registerTanshuSheet
  = G_SHEET.getSheetByName(CategoryType.registerTanshu);
const takeTanshuSheet
  = G_SHEET.getSheetByName(CategoryType.takeTanshu);
const drawLotForSchoolingSheet
  = G_SHEET.getSheetByName(CategoryType.drawLotForSchooling);
const proceduresGraduationThesesSheet
  = G_SHEET.getSheetByName(CategoryType.proceduresGraduationTheses);
const applicationForMidYearFraduationSheet
  = G_SHEET.getSheetByName(CategoryType.applicationForMidYearFraduation);

// 毎日作動させる
// Dateの演算はここでしかやらないので他で処理しなくていい
const perDayFunction = function() {
  // まず予定を取得しておく
  const postReportScheduleObj
    = getScheduleObj(postReportSheet);
  const postMediaReportScheduleObj
    = getScheduleObj(postMediaReportSheet);
  const takeTanshuScheduleObj
    = getScheduleObj(takeTanshuSheet);
  const drawLotForSchoolingScheduleObj
    = getScheduleObj(drawLotForSchoolingSheet);
  const proceduresGraduationThesesScheduleObj
    = getScheduleObj(proceduresGraduationThesesSheet);
  const applicationForMidYearFraduationScheduleObj
    = getScheduleObj(applicationForMidYearFraduationSheet);
  // リスト化しておく
  const scheduleObjList = [
    postReportScheduleObj
    ,postMediaReportScheduleObj
    ,takeTanshuScheduleObj
    ,drawLotForSchoolingScheduleObj
    ,proceduresGraduationThesesScheduleObj
    ,applicationForMidYearFraduationScheduleObj];

  // 費用については別途取得しておく
  // とはいえ、スクーリングの抽選ぐらいしか想定ができないが...
  const drawLotForSchoolingExpensesObj
    = getExpensesScheduleObj(drawLotForSchoolingSheet);
  const proceduresGraduationThesesExpensesObj
    = getExpensesScheduleObj(proceduresGraduationThesesSheet);
  const applicationForMidYearFraduationExpensesObj
    = getExpensesScheduleObj(applicationForMidYearFraduationSheet);
  // リスト化しておく
  const expensesScheduleObjList = [
    drawLotForSchoolingExpensesObj
    ,proceduresGraduationThesesExpensesObj
    ,applicationForMidYearFraduationExpensesObj];

  // 合体させる
  const allScheduleObjList = scheduleObjList.concat(expensesScheduleObjList);
  
  // 通知するべきものに対して適用する
  const today = new Date();
  allScheduleObjList.forEach(scheduleObj => {
    const sheetName = scheduleObj.sheetName;
    scheduleObj.scheduleList.forEach(schedule => {
      const starting = new Date(schedule.starting);
      const end = new Date(schedule.end);
      const expensesStarting = new Date(schedule.expensesStarting);
      const expensesEnd = new Date(schedule.expensesEnd);
      // 例外：【単修試験の試験日】は開始日2日前のみ
      if(sheetName === CategoryType.takeTanshu) {
        if(
          (starting.getTime()-(today.getTime()+DateType.ONE_DAY*2) < 0)
          && (-(DateType.ONE_DAY) < starting.getTime()-(today.getTime()+DateType.ONE_DAY*2))
        ) {
          const message = getUtilSheetEdited(sheetName, schedule, NoticeType.takeTanshu);
          postChat(SLACK_TOKEN, TARGET_CHANNEL_ID, message);
        }
      }else {
        // 予定の開始日・予定の終了日の前日・予定の終了日
        if(
          ((starting.getTime()-today.getTime() < 0) && (-(DateType.ONE_DAY) < starting.getTime()-today.getTime()))
          || ((expensesStarting.getTime()-today.getTime() < 0) && (-(DateType.ONE_DAY) < expensesStarting.getTime()-today.getTime()))
        ) {
          const message = getUtilSheetEdited(sheetName, schedule, NoticeType.startDay);
          postChat(SLACK_TOKEN, TARGET_CHANNEL_ID, message);
        }else if(
          ((end.getTime()-(today.getTime()+DateType.ONE_DAY) < 0) && (-(DateType.ONE_DAY) < end.getTime()-(today.getTime()+DateType.ONE_DAY)))
          || ((expensesEnd.getTime()-(today.getTime()+DateType.ONE_DAY) < 0) && (-(DateType.ONE_DAY) < expensesEnd.getTime()-(today.getTime()+DateType.ONE_DAY)))
        ) {
          const message = getUtilSheetEdited(sheetName, schedule, NoticeType.endBeforeOneDay);
          postChat(SLACK_TOKEN, TARGET_CHANNEL_ID, message);
        }else if(
          ((end.getTime()-today.getTime() < 0) && (-(DateType.ONE_DAY) < end.getTime()-today.getTime()))
          || (((expensesEnd.getTime()-today.getTime() < 0) && (-(DateType.ONE_DAY) < expensesEnd.getTime()-today.getTime())))
        ) {
          const message = getUtilSheetEdited(sheetName, schedule, NoticeType.endDay);
          postChat(SLACK_TOKEN, TARGET_CHANNEL_ID, message);
        }
      }
    })
  })
}

// utilシートの内容をあてはめてたオブジェクトを返す
// 通知するべき属性は引数で渡される
const getUtilSheetEdited = function(category, schedule, noticeType) {
  const today = new Date();
  const end = new Date(schedule.end);
  const expenseEnd = new Date(schedule.expensesEnd);
  // 通常の期日について
  const scheduleTitle = schedule.title;
  const scheduleStartMonth = getEditedDateObj(schedule.starting).month;
  const scheduleEnd = getEditedDateObj(schedule.end).edited;
  const remainDaysForEnd = Math.round((end.getTime() - today.getTime()) / DateType.ONE_DAY);
  // 単修が予定されていなければnull
  const monthOfRegisterTanshu = schedule.monthOfRegisterTanshu;
  // 費用について
  const expensesScheduleEnd = getEditedDateObj(schedule.expensesEnd).edited;
  const expensesRemainDaysForEnd = Math.round((expenseEnd.getTime() - today.getTime()) / DateType.ONE_DAY);
  // 単修が予定されていなければnull
  const isExpensesSchedule = schedule.expensesEnd !== '' ? true: false;

  // methods
  const thisGetUtilSheetObjValues = (rowNum) => {
    const sheetTextObj = {
      icon: '',
      starting: '',
      end: '',
      exceptional: '',
    };
    const values = utilSheet.getRange(`B${rowNum}:F${rowNum}`).getValues()[0]; 
    sheetTextObj.icon = values[0];
    sheetTextObj.starting = values[2];
    sheetTextObj.end = values[3];
    sheetTextObj.exceptional = values[4];
    return sheetTextObj;
  }
  const thisReplaceAsteriskWithTerm = (text, term) => {
    // 大文字のアスタリスクなので注意
    const searchTerm = '＊';
    return text.replace(searchTerm, term);
  }

  // それぞれの場合に対してメッセージを作成する
  // 【リポートの提出】
  if(category === CategoryType.postReport){
    const rowNum = 2;
    const utilSheetObj = thisGetUtilSheetObjValues(rowNum);
    const header = `${utilSheetObj.icon}${category}${utilSheetObj.icon}`
    let bodyReplacedAsterisk = '';
    if(noticeType === NoticeType.startDay) {
      if(monthOfRegisterTanshu === null) {
        // 【あと＊日】
        // リポート提出期間になりました！
        // なお、今回は単修試験を直後に控えていません。
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.exceptional, remainDaysForEnd);
        // 期限は＊です。
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleEnd);
      }else {
        // 【あと＊日】
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.starting, remainDaysForEnd);
        // ＊月試験のためのリポート提出期間になりました！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, monthOfRegisterTanshu);
        // 期限は＊です。
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleEnd);
      }
    }else if(noticeType === NoticeType.endBeforeOneDay) {
      // ＊はリポート提出期限です！
      // ご提出をお忘れなく！
      bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.end, RemindType.oneDayBefore);
    }else if(noticeType === NoticeType.endDay) {
      // ＊はリポート提出期限です！
      // ご提出をお忘れなく！
      bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.end, RemindType.today);
    }
    // 単修試験の登録と被っていないか判定
    if(monthOfRegisterTanshu === null) {
      const message = header + '\n' + bodyReplacedAsterisk;
      return message;
    } else {
      const registerTanshuRowNum = 4;
      const registerTanshuUtilSheetObj 
        = thisGetUtilSheetObjValues(registerTanshuRowNum);
      const registerTanshuHeader 
        = `${registerTanshuUtilSheetObj.icon}${CategoryType.registerTanshu}${registerTanshuUtilSheetObj.icon}`
      // 試験登録もお忘れなく！
      const registerTanshubody
        = registerTanshuUtilSheetObj.starting;
      const message 
        = header + '\n' + bodyReplacedAsterisk
        + '\n' + registerTanshuHeader 
        + '\n' + registerTanshubody;
      return message;
    }
  // 【Mスクのリポート提出】
  }else if(category === CategoryType.postMediaReport){
    const rowNum = 3;
    const utilSheetObj = thisGetUtilSheetObjValues(rowNum);
    const header = `${utilSheetObj.icon}${category}${utilSheetObj.icon}`
    let bodyReplacedAsterisk = '';
    if(noticeType === NoticeType.startDay) {
      // 【あと＊日】
      bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.starting, remainDaysForEnd);
      // ＊のためのリポート提出期間になりました！
      bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      // 期限は＊です。
      bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.starting, scheduleEnd);
    }else if(noticeType === NoticeType.endBeforeOneDay) {
      // ＊は
      bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.end, RemindType.oneDayBefore);
      // ＊のためのリポート提出期限です！
      bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
    }else if(noticeType === NoticeType.endDay) {
      // ＊は
      bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.end, RemindType.today);
      // ＊のためのリポート提出期限です！
      bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
    }
    return header + '\n' + bodyReplacedAsterisk;
  }else if(category === CategoryType.takeTanshu){
    // 例外：【単修試験の試験日】は開始日2日前のみ
    const rowNum = 5;
    const utilSheetObj = thisGetUtilSheetObjValues(rowNum);
    const header = `${utilSheetObj.icon}単修試験${utilSheetObj.icon}`
    if(noticeType == 0) {
      // 【いよいよ日曜日】
      // 週末は＊月単修試験の実施日です！
      // 受験票など、持ち物をお忘れなく！
      bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.starting, scheduleStartMonth);
    }
    return header + '\n' + bodyReplacedAsterisk;
  // 【スクーリングの抽選】
  }else if(category === CategoryType.drawLotForSchooling){
    const rowNum = 6;
    const utilSheetObj = thisGetUtilSheetObjValues(rowNum);
    const header = `${utilSheetObj.icon}${category}${utilSheetObj.icon}`
    let bodyReplacedAsterisk = '';
    if(isExpensesSchedule) {
      // 【費用の振込】
      const expensesRowNum = 7;
      const expensesSheetTextObj = thisGetUtilSheetObjValues(expensesRowNum);
      if(noticeType === NoticeType.startDay) {
        // 【あと＊日】
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(expensesSheetTextObj.starting, expensesRemainDaysForEnd);
        // ＊のための費用振込み期間になりました！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
        // 期限は＊です。
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, expensesScheduleEnd);
      }else if(noticeType === NoticeType.endBeforeOneDay) {
        // ＊は
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(expensesSheetTextObj.end, RemindType.oneDayBefore);
        // ＊のための費用振込み期限です！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      }else if(noticeType === NoticeType.endDay) {
        // ＊は
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(expensesSheetTextObj.end, RemindType.today);
        // ＊のための費用振込み期限です！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      }
    }else {
      if(noticeType === NoticeType.startDay) {
        // 【あと＊日】
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.starting, remainDaysForEnd);
        // ＊スクーリングの抽選期間になりました！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
        // 期限は＊です。
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleEnd);
      }else if(noticeType === NoticeType.endBeforeOneDay) {
        // ＊は
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.end, RemindType.oneDayBefore);
        // ＊スクーリングの抽選期限です！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      }else if(noticeType === NoticeType.endDay) {
        // ＊は
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.end, RemindType.today);
        // ＊スクーリングの抽選期限です！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      }
    }
    return header + '\n' + bodyReplacedAsterisk;
  // 【卒業論文関連の手続き】
  // 未定義
  }else if(category === CategoryType.proceduresGraduationTheses){
    const rowNum = 8;
    const utilSheetObj = thisGetUtilSheetObjValues(rowNum);
    const header = `${utilSheetObj.icon}${category}${utilSheetObj.icon}`
    let bodyReplacedAsterisk = '';
    if(isExpensesSchedule) {
      // 【費用の振込】
      const expensesRowNum = 7;
      const expensesSheetTextObj = thisGetUtilSheetObjValues(expensesRowNum);
      if(noticeType === NoticeType.startDay) {
        // 【あと＊日】
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(expensesSheetTextObj.starting, expensesRemainDaysForEnd);
        // ＊のための費用振込み期間になりました！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
        // 期限は＊です。
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, expensesScheduleEnd);
      }else if(noticeType === NoticeType.endBeforeOneDay) {
        // ＊は
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(expensesSheetTextObj.end, RemindType.oneDayBefore);
        // ＊のための費用振込み期限です！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      }else if(noticeType === NoticeType.endDay) {
        // ＊は
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(expensesSheetTextObj.end, RemindType.today);
        // ＊のための費用振込み期限です！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      }
    }else {
      if(noticeType === NoticeType.startDay) {
        // 【あと＊日】
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.starting, remainDaysForEnd);
        // ＊のためのリポート提出期間になりました！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
        // 期限は＊です。
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleEnd);
      }else if(noticeType === NoticeType.endBeforeOneDay) {
        // ＊は
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.end, RemindType.oneDayBefore);
        // ＊のための手続き期限です！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      }else if(noticeType === NoticeType.endDay) {
        // ＊は
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.end, RemindType.today);
        // ＊のための手続き期限です！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      }
    }
    return header + '\n' + bodyReplacedAsterisk;
  // 【年度途中卒業の申請】
  // 未定義
  }else if(category === CategoryType.applicationForMidYearFraduation){
    const rowNum = 9;
    const utilSheetObj = thisGetUtilSheetObjValues(rowNum);
    const header = `${utilSheetObj.icon}${category}${utilSheetObj.icon}`
    let bodyReplacedAsterisk = '';
    if(isExpensesSchedule) {
      // 【費用の振込】
      const expensesRowNum = 7;
      const expensesSheetTextObj = thisGetUtilSheetObjValues(expensesRowNum);
      if(noticeType === NoticeType.startDay) {
        // 【あと＊日】
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(expensesSheetTextObj.starting, expensesRemainDaysForEnd);
        // ＊のための費用振込み期間になりました！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
        // 期限は＊です。
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, expensesScheduleEnd);
      }else if(noticeType === NoticeType.endBeforeOneDay) {
        // ＊は
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(expensesSheetTextObj.end, RemindType.oneDayBefore);
        // ＊のための費用振込み期限です！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      }else if(noticeType === NoticeType.endDay) {
        // ＊は
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(expensesSheetTextObj.end, RemindType.today);
        // ＊のための費用振込み期限です！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      }
    }else {
      if(noticeType === NoticeType.startDay) {
        // 【あと＊日ぐらい】
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.starting, remainDaysForEnd);
        // そろそろ＊月卒業のための申込が可能になります！
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(bodyReplacedAsterisk, scheduleTitle);
      }else if(noticeType === NoticeType.endBeforeOneDay) {
        // そろそろ＊卒業のための申込の期限が終了になります！
        // 詳細はHPや法政通信を確認してください。
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.end, scheduleTitle);
      }else if(noticeType === NoticeType.endDay) {
        // そろそろ＊卒業のための申込の期限が終了になります！
        // 詳細はHPや法政通信を確認してください。
        bodyReplacedAsterisk = thisReplaceAsteriskWithTerm(utilSheetObj.end, scheduleTitle);
      }
    }
    return header + '\n' + bodyReplacedAsterisk;
  }
}

// 各カテゴリーのシートから予定を取得する
const getScheduleObj = function(sheet) {
  const lastRow = sheet.getLastRow();
  const scheduleList = new Array();
  const title = sheet.getRange(`A2:A${lastRow}`).getValues(); 
  const starting = sheet.getRange(`B2:B${lastRow}`).getValues(); 
  const end = sheet.getRange(`C2:C${lastRow}`).getValues(); 
  for(let i=0; i<starting.length; i++){
    const obj = {
      title: '',
      starting: '',
      end: '',
      expensesStarting: '',
      expensesEnd: '',
      monthOfRegisterTanshu: null,
    };
    obj.title = title[i][0];
    obj.starting = starting[i][0];
    obj.end = end[i][0];
    // リポートの提出と単修試験の登録が一致しているか判定する
    if(sheet.getSheetName() === CategoryType.postReport) {
      const tanshuArr = [4, 5, 6, null, 7, null, null, 10, 11, 12, 1, null];
      obj.monthOfRegisterTanshu = tanshuArr[i];
    }
    scheduleList.push(obj);
  }
  return {
    sheetName: sheet.getSheetName(),
    scheduleList: scheduleList,
  };
}

// 各カテゴリーのシートから費用についての予定を取得する
const getExpensesScheduleObj = function(sheet) {
  const lastRow = sheet.getLastRow();
  const expensesScheduleList = new Array();
  const title = sheet.getRange(`A2:A${lastRow}`).getValues(); 
  const starting = sheet.getRange(`E2:E${lastRow}`).getValues(); 
  const end = sheet.getRange(`F2:F${lastRow}`).getValues(); 
  for(let i=0; i<starting.length; i++){
    const obj = {
      title: '',
      starting: '',
      end: '',
      expensesStarting: '',
      expensesEnd: '',
      monthOfRegisterTanshu: null,
    };
    obj.title = title[i][0];
    obj.expensesStarting = starting[i][0];
    obj.expensesEnd = end[i][0];
    expensesScheduleList.push(obj);
  }
  return {
    sheetName: sheet.getSheetName(),
    scheduleList: expensesScheduleList,
  };
}

// Date型のデータを良い感じに整形して返す
const getEditedDateObj = function(data) {
  // 予定の開始時間を取得
  const dayArray = ['日', '月', '火', '水', '木', '金', '土']
  const typedDate = new Date(data);
  const month = typedDate.getMonth()+1;
  const date = typedDate.getDate();
  const day = typedDate.getDay();
  const edited = 
    month + '月' + date + '日' + `(${dayArray[day]})`;
  return {
    'month': month,
    'date': date,
    'edited': edited,
  }
}

const getChannelId = function(token, channelName) {
  let channelId = '';
  const result = getChannelList(token);
  if (result.ok) {
    for (let channel of result.channels) {
      if (channel.name == channelName) {
        channelId = channel.id;
        break;
      }
    }
  }
  return channelId;
}

const getChannelList = function(token) {
  const methodUrl = 'https://slack.com/api/conversations.list';
  const payload = {
    'token': token,
    'types': 'private_channel'
    // 'is_private': false,
  };
  const option = {
    'method': 'POST',
    'payload': payload,
  };
  const response = UrlFetchApp.fetch(methodUrl, option);
  return JSON.parse(response)
}

const postChat = function(token, channel, message) {
  const methodUrl = 'https://slack.com/api/chat.postMessage';
  const payload = {
    'token': token,
    'channel': channel,
    'text': message,
  };
  const option = {
    'method': 'POST',
    'payload': payload,
  };
  const response = UrlFetchApp.fetch(methodUrl, option);
  return JSON.parse(response)
}