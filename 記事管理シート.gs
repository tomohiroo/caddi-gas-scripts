function notifyStatus() {
  var mySheet = SpreadsheetApp.getActiveSheet();
  var myCell = mySheet.getActiveCell();
  if (myCell.getColumn() != 2) {
    return;
  }
  // ステータスが更新されたら
  var writerName = mySheet.getSheetName();
  var keyword = myCell.offset(0, -1).getValue();
  var status = myCell.getValue();
  if (status == "" || writerName == "集計") {
    return;
  } else if (writerName == "tomohiro_ueda" || writerName == "sara_yamaguchi") {
    var mentionTo = mentionToCaddiInCharge(status, writerName);
    var channel = "#seo-writing";
  } else {
    var mentionTo = statusText2Mention(status, writerName);
    var channel = "#seo-writing-team";
  }
  var message =
    writerName +
    "様の「" +
    keyword +
    "」の記事のステータスが" +
    status +
    "になりました";
  slackPost(message, mentionTo, channel);
}

function slackPost(message, mentionTo, channel) {
  var url =
    "https://hooks.slack.com/services/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";
  var username = "記事ステータス更新通知bot";
  var method = "post";
  var icon_emoji = ":ghost:";
  var payload = JSON.stringify({
    channel: channel,
    text: mentionTo + "\n" + message,
    username: username,
    icon_emoji: icon_emoji
  });
  var params = {
    method: method,
    payload: payload
  };
  var response = UrlFetchApp.fetch(url, params);
}

function statusText2Mention(status, writerName) {
  var statusMapping = {
    kwボツ: mentionToCaddiMember(),
    骨子作成中: userName2Mention(writerName),
    骨子レビュー待ち: mentionToCaddiMember(),
    記事作成中: userName2Mention(writerName),
    記事レビュー待ち: mentionToCaddiMember(),
    記事レビュー済: userName2Mention(writerName),
    納品完了: mentionToCaddiMember(),
    記事完成: "",
    リリース済: userName2Mention(writerName)
  };
  return statusMapping[status];
}

function mentionToCaddiInCharge(status, writerName) {
  var writer = writerName == "tomohiro_ueda" ? "tomohiro_ueda" : "sara_yamaguchi";
  var reviewer = writerName == "tomohiro_ueda" ? "sara_yamaguchi" : "tomohiro_ueda";
  var statusMapping = {
    kwボツ: userName2Mention(reviewer),
    骨子作成中: userName2Mention(writerName),
    骨子レビュー待ち: userName2Mention(reviewer),
    記事作成中: userName2Mention(writer),
    記事レビュー待ち: userName2Mention(reviewer),
    記事レビュー済: userName2Mention(writer),
    納品完了: userName2Mention(writer),
    記事完成: "",
    リリース済: mentionToCaddiMember()
  };
  return statusMapping[status];
}

function userName2Mention(name) {
  var userIdMapping = {
    miyabi_mito: "UKDF6830U",
    saori_kawashima: "UKE5B8A69",
    tomohiro_ueda: "UG0SMH45T",
    sara_yamaguchi: "UFLFKGN3E",
    yuta_yamamoto: "UG0SMH45T"
  };
  return "<@" + userIdMapping[name] + ">";
}

function mentionToCaddiMember() {
  return "<@UG0SMH45T> <@UFLFKGN3E>";
}




function syncronizeKpi() {
  var articleManegementSheets = SpreadsheetApp.openById(
    "1cTwrslRYnfwprJETR1GsIWhsjFG2cMkMRKXQDHE3Rbo"
  );
  var aggregateSheet = articleManegementSheets.getSheetByName("集計");

  var completeArticleNumbers = aggregateSheet.getRange("M:M").getValues();
  var inspectedArticleNumbers = aggregateSheet.getRange("N:N").getValues();
  var releasedArticleNumbers = aggregateSheet.getRange("O:O").getValues();

  var completeArticleNumber = completeArticleNumbers
    .slice(1)
    .reduce(function(prev, current) {
      return parseInt(prev) + parseInt(current);
    });
  var inspectedArticleNumber = inspectedArticleNumbers
    .slice(1)
    .reduce(function(prev, current) {
      return parseInt(prev) + parseInt(current);
    });
  var releasedArticleNumber = releasedArticleNumbers
    .slice(1)
    .reduce(function(prev, current) {
      return parseInt(prev) + parseInt(current);
    });

  var kpiSheets = SpreadsheetApp.openById(
    "1dt5hC6eSLeFG74CT1yKSgjuWBOkan23gjbVhOxUkw30"
  );
  var kpiSheet = kpiSheets.getSheetByName("全体KPI");
  var todayKpiRowNum =
    kpiSheet
      .getRange("5:5")
      .getValues()[0]
      .slice(4)
      .filter(function(v) {
        return v != "";
      }).length + 5;
  kpiSheet.getRange(5, todayKpiRowNum).setValue(completeArticleNumber);
  kpiSheet.getRange(7, todayKpiRowNum).setValue(inspectedArticleNumber);
  kpiSheet.getRange(9, todayKpiRowNum).setValue(releasedArticleNumber);
}

function trackCwWriterProgress() {
  var articleManegementSheets = SpreadsheetApp.openById(
    "1cTwrslRYnfwprJETR1GsIWhsjFG2cMkMRKXQDHE3Rbo"
  );
  var aggregateSheet = articleManegementSheets.getSheetByName("集計");

  var writerNames = aggregateSheet.getRange("A:A").getValues();
  var writerStatuses = aggregateSheet.getRange("B:B").getValues();
  var cwWriterNames = [];
  var cwWriterColumnNumbers = [];
  writerStatuses.forEach(function(status, idx) {
    if (status == "CW") {
      cwWriterNames = cwWriterNames.concat(writerNames[idx]);
      cwWriterColumnNumbers = cwWriterColumnNumbers.concat(idx + 1);
    }
  });

  var kpiSheets = SpreadsheetApp.openById(
    "1dt5hC6eSLeFG74CT1yKSgjuWBOkan23gjbVhOxUkw30"
  );
  cwWriterNames.forEach(function(name, idx) {
    kpiSheet = kpiSheets.getSheetByName(name);
    kpi = aggregateSheet
      .getRange(cwWriterColumnNumbers[idx], 13, 1, 3)
      .getValues()[0];
    todayKpiRowNum =
      kpiSheet
        .getRange("5:5")
        .getValues()[0]
        .slice(4)
        .filter(function(v) {
          return v !== "";
        }).length + 5;
    kpiSheet.getRange(5, todayKpiRowNum).setValue(kpi[0]);
    kpiSheet.getRange(7, todayKpiRowNum).setValue(kpi[1]);
    kpiSheet.getRange(9, todayKpiRowNum).setValue(kpi[2]);
  });
}
