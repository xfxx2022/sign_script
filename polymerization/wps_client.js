/*
    name: "WPS客户端版打卡"
    cron: 30 0 9 * * *
    脚本兼容: 金山文档（1.0），金山文档（2.0）
    更新时间：20260711
    环境变量名：无
    环境变量值：无
    备注：WPS客户端版/会员时长打卡（当前可用端点 zt.wps.cn/2018/clock_in/api/clock_in）。
          配合“wps”分配置表使用（cookie 即 wps_sid，作为 sid 请求头使用）。
          “wps”分配置表列：A=cookie，B=是否执行，C=账号名称，
          D=转存PPT(稻壳版用)，E=是否渠道1打卡，F=是否渠道2打卡(本脚本用)，G=预留列(未使用)
          说明：本脚本读取列F（是否渠道2打卡），为空或“是”时执行，填“否”则跳过。
          原 vipapi.wps.cn/wps_clock/v2 端点已于早年停用且需未公开签名，故改用当前可用端点；
          该端点使用 sid 请求头（wps_sid 值），无需额外的 Signature，列G为预留列。
          若返回“前一天未报名”，会自动报名(sign_up)后重试一次。
*/

var sheetNameSubConfig = "wps"; // 分配置表名称（共享的分表，subConfigWps 的键）
let sheetNameSubConfig2 = "wps_client"; // 在 CONFIG 表中的唯一名称，用于读取个性化配置与写入消息队列
var pushHeader = "【WPS客户端版】"; // 推送头，给自己看的
var sheetNameConfig = "CONFIG"; // 总配置表
var sheetNamePush = "PUSH"; // 推送表名称
var sheetNameEmail = "EMAIL"; // 邮箱表
var flagSubConfig = 0; // 激活分配置工作表标志
var flagConfig = 0; // 激活主配置工作表标志
var flagPush = 0; // 激活推送工作表标志
var line = 21; // 指定读取从第2行到第line行的内容
var message = ""; // 待发送的消息
var messageArray = [];  // 待发送的消息数据，每个元素都是某个账号的消息。目的是将不同用户消息分离，方便个性化消息配置
var messageOnlyError = 0; // 0为只推送失败消息，1则为推送成功消息。
var messageNickname = 0; // 1为推送位置标识（昵称/单元格Ax（昵称为空时）），0为不推送位置标识
var messageHeader = []; // 存放每个消息的头部，如：单元格A3。目的是分离附加消息和执行结果消息
var messagePushHeader = pushHeader; // 存放在总消息的头部，默认是pushHeader,如：【xxxx】
var version = 1 // 版本类型，自动识别并适配。默认为airscript 1.0，否则为2.0（Beta）
var separator = "##########MOKU##########" // 分割符，分割消息。可用于PUSH.js灵活推送
var maxMessageLength = 400;  // 设置最大长度，超过这个长度则分片发送
var messageDistance = 100; // 消息距离，用于匹配100字符内最近的行

var jsonPush = [
  { name: "bark", key: "xxxxxx", flag: "0" },
  { name: "pushplus", key: "xxxxxx", flag: "0" },
  { name: "ServerChan", key: "xxxxxx", flag: "0" },
  { name: "email", key: "xxxxxx", flag: "0" },
  { name: "dingtalk", key: "xxxxxx", flag: "0" },
  { name: "discord", key: "xxxxxx", flag: "0" },
]; // 推送数据，flag=1则推送
var jsonEmail = {
  server: "",
  port: "",
  sender: "",
  authorizationCode: "",
}; // 有效邮箱配置

// =================青龙适配开始===================

qlSwitch = 0

// =================青龙适配结束===================

// =================金山适配开始===================
// airscript检测版本
function checkVesion(){
  try{
    let temp = Application.Range("A1").Text;
    Application.Range("A1").Value  = temp
    console.log("😶‍🌫️ 检测到当前airscript版本为1.0，进行1.0适配")
  }catch{
    console.log("😶‍🌫️ 检测到当前airscript版本为2.0，进行2.0适配")
    version = 2
  }
}

// 推送相关
// 获取时间
function getDate(){
  let currentDate = new Date();
  currentDate = currentDate.getFullYear() + '/' + (currentDate.getMonth() + 1).toString() + '/' + currentDate.getDate().toString();
  return currentDate
}

// 将消息写入CONFIG表中作为消息队列，之后统一发送
function writeMessageQueue(message){
  // 当天时间
  let todayDate = getDate()
  flagConfig = ActivateSheet(sheetNameConfig); // 激活主配置表
  // 主配置工作表存在
  if (flagConfig == 1) {
    console.log("✨ 开始将结果写入主配置表");
    for (let i = 2; i <= 100; i++) {
      // 找到指定的表行
      if(Application.Range("A" + (i + 2)).Value == sheetNameSubConfig2){
        // 写入更新的时间
        Application.Range("F" + (i + 2)).Value = todayDate
        // 写入消息
        Application.Range("G" + (i + 2)).Value = message
        console.log("✨ 写入结果完成");
        break;
      }
    }
  }

}

// 总推送
function push(message) {
  writeMessageQueue(message)  // 将消息写入CONFIG表中
  // if (message != "") {
  //   // message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
  //   let length = jsonPush.length;
  //   let name;
  //   let key;
  //   for (let i = 0; i < length; i++) {
  //     if (jsonPush[i].flag == 1) {
  //       name = jsonPush[i].name;
  //       key = jsonPush[i].key;
  //       if (name == "bark") {
  //         bark(message, key);
  //       } else if (name == "pushplus") {
  //         pushplus(message, key);
  //       } else if (name == "ServerChan") {
  //         serverchan(message, key);
  //       } else if (name == "email") {
  //         email(message);
  //       } else if (name == "dingtalk") {
  //         dingtalk(message, key);
  //       } else if (name == "discord") {
  //         discord(message, key);
  //       }
  //     }
  //   }
  // } else {
  //   console.log("🍳 消息为空不推送");
  // }
}

// 推送bark消息
function bark(message, key) {
    if (key != "") {
      message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
      message = encodeURIComponent(message)
      BARK_ICON = "https://s21.ax1x.com/2024/06/23/pkrUkfe.png"
    let url = "https://api.day.app/" + key + "/" + message + "/" + "?icon=" + BARK_ICON;
    // 若需要修改推送的分组，则将上面一行改为如下的形式
    // let url = 'https://api.day.app/' + bark_id + "/" + message + "?group=分组名";
    let resp = HTTP.get(url, {
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
    });
    sleep(5000);
    }
}

// 推送pushplus消息
function pushplus(message, key) {
  if (key != "") {
      message = encodeURIComponent(message)
    // url = "http://www.pushplus.plus/send?token=" + key + "&content=" + message;
    url = "http://www.pushplus.plus/send?token=" + key + "&content=" + message + "&title=" + pushHeader;  // 增加标题
    let resp = HTTP.fetch(url, {
      method: "get",
    });
    sleep(5000);
  }
}

// 推送serverchan消息
function serverchan(message, key) {
  if (key != "") {
    url =
      "https://sctapi.ftqq.com/" +
      key +
      ".send" +
      "?title=" + messagePushHeader +
      "&desp=" +
      message;
    let resp = HTTP.fetch(url, {
      method: "get",
    });
    sleep(5000);
  }
}

// email邮箱推送
function email(message) {
  var myDate = new Date(); // 创建一个表示当前时间的 Date 对象
  var data_time = myDate.toLocaleDateString(); // 获取当前日期的字符串表示
  let server = jsonEmail.server;
  let port = parseInt(jsonEmail.port); // 转成整形
  let sender = jsonEmail.sender;
  let authorizationCode = jsonEmail.authorizationCode;

  let mailer;
  mailer = SMTP.login({
    host: server,
    port: port,
    username: sender,
    password: authorizationCode,
    secure: true,
  });
  mailer.send({
    from: pushHeader + "<" + sender + ">",
    to: sender,
    subject: pushHeader + " - " + data_time,
    text: message,
  });
  // console.log("🍳 已发送邮件至：" + sender);
  console.log("🍳 已发送邮件");
  sleep(5000);
}

// 邮箱配置
function emailConfig() {
  console.log("🍳 开始读取邮箱配置");
  let length = jsonPush.length; // 因为此json数据可无序，因此需要遍历
  let name;
  for (let i = 0; i < length; i++) {
    name = jsonPush[i].name;
    if (name == "email") {
      if (jsonPush[i].flag == 1) {
        let flag = ActivateSheet(sheetNameEmail); // 激活邮箱表
        // 邮箱表存在
        // var email = {
        //   'email':'', 'port':'', 'sender':'', 'authorizationCode':''
        // } // 有效配置
        if (flag == 1) {
          console.log("🍳 开始读取邮箱表");
          for (let i = 2; i <= 2; i++) {
            // 从工作表中读取推送数据
            jsonEmail.server = Application.Range("A" + i).Text;
            jsonEmail.port = Application.Range("B" + i).Text;
            jsonEmail.sender = Application.Range("C" + i).Text;
            jsonEmail.authorizationCode = Application.Range("D" + i).Text;
            if (Application.Range("A" + i).Text == "") {
              // 如果为空行，则提前结束读取
              break;
            }
          }
          // console.log(jsonEmail)
        }
        break;
      }
    }
  }
}

// 推送钉钉机器人
function dingtalk(message, key) {
  message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
  let url = "https://oapi.dingtalk.com/robot/send?access_token=" + key;
  let resp = HTTP.post(url, { msgtype: "text", text: { content: message } });
  // console.log(resp.text())
  sleep(5000);
}

// 推送Discord机器人
function discord(message, key) {
  message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
  let url = key;
  let resp = HTTP.post(url, { content: message });
  //console.log(resp.text())
  sleep(5000);
}

// =================金山适配结束===================
// =================共用开始===================
// main()  // 入口

// function main(){
  checkVesion() // 版本检测，以进行不同版本的适配

  flagConfig = ActivateSheet(sheetNameConfig); // 激活推送表
  // 主配置工作表存在
  if (flagConfig == 1) {
    console.log("🍳 开始读取主配置表");
    let name; // 名称
    let onlyError;
    let nickname;
    for (let i = 2; i <= 100; i++) {
      // 从工作表中读取推送数据
      name = Application.Range("A" + i).Text;
      onlyError = Application.Range("C" + i).Text;
      nickname = Application.Range("D" + i).Text;
      if (name == "") {
        // 如果为空行，则提前结束读取
        break; // 提前退出，提高效率
      }
      if (name == sheetNameSubConfig2) {
        if (onlyError == "是") {
          messageOnlyError = 1;
          console.log("🍳 只推送错误消息");
        }

        if (nickname == "是") {
          messageNickname = 1;
          console.log("🍳 单元格用昵称替代");
        }

        break; // 提前退出，提高效率
      }
    }
  }

  flagPush = ActivateSheet(sheetNamePush); // 激活推送表
  // 推送工作表存在
  if (flagPush == 1) {
    console.log("🍳 开始读取推送工作表");
    let pushName; // 推送类型
    let pushKey;
    let pushFlag; // 是否推送标志
    for (let i = 2; i <= line; i++) {
      // 从工作表中读取推送数据
      pushName = Application.Range("A" + i).Text;
      pushKey = Application.Range("B" + i).Text;
      pushFlag = Application.Range("C" + i).Text;
      if (pushName == "") {
        // 如果为空行，则提前结束读取
        break;
      }
      jsonPushHandle(pushName, pushFlag, pushKey);
    }
    // console.log(jsonPush)
  }

  // 邮箱配置函数
  emailConfig();

  flagSubConfig = ActivateSheet(sheetNameSubConfig); // 激活分配置表
  if (flagSubConfig == 1) {
    console.log("🍳 开始读取分配置表");

      if(qlSwitch != 1){  // 金山文档
          for (let i = 2; i <= line; i++) {
              var cookie = Application.Range("A" + i).Text;
              var exec = Application.Range("B" + i).Text;
              if (cookie == "") {
                  // 如果为空行，则提前结束读取
                  break;
              }
              if (exec == "是") {
                  execHandle(cookie, i);
              }
          }
          message = messageMerge()// 将消息数组融合为一条总消息
          push(message); // 推送消息
      }else{
          for (let i = 2; i <= line; i++) {
              var cookie = Application.Range("A" + i).Text;
              var exec = Application.Range("B" + i).Text;
              if (cookie == "") {
                  // 如果为空行，则提前结束读取
                  break;
              }
              if (exec == "是") {
                  console.log("🧑 开始执行用户：" + "1" )
                  execHandle(cookie, i);
                  break;  // 只取一个
              }
          }
      }

  }

// }

// 激活工作表函数
function ActivateSheet(sheetName) {
    let flag = 0;
    try {
      // 激活工作表
      let sheet = Application.Sheets.Item(sheetName);
      sheet.Activate();
      console.log("🥚 激活工作表：" + sheet.Name);
      flag = 1;
    } catch {
      flag = 0;
      console.log("🍳 无法激活工作表，工作表可能不存在");
    }
    return flag;
}

// 对推送数据进行处理
function jsonPushHandle(pushName, pushFlag, pushKey) {
  let length = jsonPush.length;
  for (let i = 0; i < length; i++) {
    if (jsonPush[i].name == pushName) {
      if (pushFlag == "是") {
        jsonPush[i].flag = 1;
        jsonPush[i].key = pushKey;
      }
    }
  }
}

// 将消息数组融合为一条总消息
function messageMerge(){
    // console.log(messageArray)
    let message = ""
  for(i=0; i<messageArray.length; i++){
    if(messageArray[i] != "" && messageArray[i] != null)
    {
      message += "\n" + messageHeader[i] + messageArray[i] + ""; // 加上推送头
    }
  }
  if(message != "")
  {
    console.log("✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨")
    console.log(message + "\n")  // 打印总消息
    console.log("✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨")
  }
  return message
}

function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d; );
}

// 获取sign，返回小写
function getsign(data) {
    var sign = Crypto.createHash("md5")
        .update(data, "utf8")
        .digest("hex")
        // .toUpperCase() // 大写
        .toString();
    return sign;
}

// =================共用结束===================

// 具体的执行函数
function execHandle(cookie, pos) {
  let messageSuccess = "";
  let messageFail = "";
  let messageName = "";
  // 推送昵称或单元格，还是不推送位置标识
  if (messageNickname == 1) {
    // 推送昵称或单元格
    messageName = Application.Range("C" + pos).Text;
    if(messageName == "")
    {
      messageName = "单元格A" + pos + "";
    }
  }

  posLabel = pos-2 ;  // 存放下标，从0开始
  messageHeader[posLabel] = "👨‍🚀 " + messageName

  try {
    // 是否启用本渠道（渠道2 = 客户端版打卡）。列F为空或“是”则执行，“否”则跳过。
    let channel2 = Application.Range("F" + pos).Text;
    if(channel2 != "" && channel2 != "是"){
      console.log("📢 " + messageName + " 未启用渠道2打卡，跳过");
      messageArray[posLabel] = "";
      return;
    }

    // 会员时长打卡：使用 sid 请求头（wps_sid 值）。
    // 走 ?member=wps 答题流程——该路径用答题代替裸端点的微信验证码墙，可自动化。
    let url = "https://zt.wps.cn/2018/clock_in/api/clock_in?member=wps";
    let signUpUrl = "https://zt.wps.cn/2018/clock_in/api/sign_up";
    let getQuestionUrl = "https://zt.wps.cn/2018/clock_in/api/get_question?member=wps";
    let answerUrl = "https://zt.wps.cn/2018/clock_in/api/answer?member=wps";
    headers = {
      "sid": cookie,
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586"
    };
    // AirScript 2.0 下 POST 必须显式声明 Content-Type，否则被 ZLB 以 400 拒绝
    postHeaders = {
      "sid": cookie,
      "Content-Type": "application/x-www-form-urlencoded",
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586"
    };

    // 发起一次打卡请求，返回解析后的响应对象或错误标记
    function clockIn() {
      let resp = HTTP.get(url, { headers: headers });
      if (resp.status == 200) {
        try {
          return resp.json();
        } catch {
          console.log(resp.text());
          return { __parseError: true };
        }
      } else {
        console.log(resp.text());
        return { __httpError: resp.status };
      }
    }

    function clockInQuestion() {
      let resp = HTTP.get(getQuestionUrl, { headers: headers });
      if (resp.status == 200) {
        try {
          return resp.json();
        } catch {
          console.log(resp.text());
          return null;
        }
      }
      console.log(resp.text());
      return null;
    }

    // 解析打卡结果：成功返回“🎉...”，其它返回 null（需进一步判断）
    function handleResult(r) {
      if (r.__httpError != undefined) {
        return "❌ 打卡失败：HTTP " + r.__httpError + "\n";
      }
      if (r.__parseError) {
        return "❌ 打卡失败：响应解析错误\n";
      }
      let msg = r["msg"] || "";
      let result = r["result"] || "";
      if (result == "ok" || msg == "已打卡" || msg == "ok") {
        return "🎉 打卡成功\n";
      }
      return null;
    }

    // 答题流程：跳过多选题，用固定答案集匹配单选题作答；答错则逐选项重试
    function doQuiz() {
      let q = clockInQuestion();
      let guard = 0;
      while (q != null && q["data"] && q["data"]["multi_select"] == 1 && guard < 10) {
        sleep(800);
        q = clockInQuestion();
        guard++;
      }
      if (q == null || !q["data"]) return true; // 无需答题则视为成功
      let options = q["data"]["options"] || [];
      // WPS 会员特权类固定答案集（社区 dailycheckin 验证）
      let answerSet = {
        "WPS会员全文检索": 1, "100G": 1, "WPS会员数据恢复": 1, "WPS会员PDF转doc": 1,
        "WPS会员PDF转图片": 1, "WPS图片转PDF插件": 1, "金山PDF转WORD": 1, "WPS会员拍照转文字": 1,
        "使用WPS会员修复": 1, "WPS全文检索功能": 1, "有,且无限次": 1, "文档修复": 1
      };
      let answerId = 3; // 默认第3项（选项索引从1开始）
      for (let i = 0; i < options.length; i++) {
        if (answerSet[options[i]] == 1) { answerId = i + 1; break; }
      }
      let resp = HTTP.post(answerUrl, { headers: postHeaders, data: { answer: answerId } });
      if (resp.status != 200) { console.log(resp.text()); return false; }
      try {
        let ar = resp.json();
        if (ar["result"] == "ok" || ar["msg"] == "ok") return true;
        if (ar["msg"] == "wrong answer") {
          // 逐选项重试
          for (let i = 1; i <= options.length; i++) {
            let r2 = HTTP.post(answerUrl, { headers: postHeaders, data: { answer: i } });
            if (r2.status == 200) {
              try {
                let ar2 = r2.json();
                if (ar2["result"] == "ok" || ar2["msg"] == "ok") return true;
              } catch {}
            }
          }
        }
      } catch {
        console.log(resp.text());
      }
      return false;
    }

    // 主流程封装：处理报名/答题/验证码兜底，返回最终消息串
    function tryClockIn() {
      let r = clockIn();
      let h = handleResult(r);
      if (h != null) return h;            // 成功或硬失败
      let msg = r["msg"] || "";
      if (msg == "前一天未报名") {
        console.log("📢 前一天未报名，尝试自动报名");
        try { HTTP.post(signUpUrl, { headers: postHeaders }); } catch {}
        sleep(1500);
        doQuiz();                          // 报名后可能触发新题
        sleep(1200);
        let r2 = clockIn();
        let h2 = handleResult(r2);
        if (h2 != null && h2.indexOf("🎉") >= 0) return "🎉 打卡成功（已自动报名）\n";
        return h2 != null ? h2 : "❌ 打卡失败：" + (r2["msg"] || "未知错误") + "\n";
      }
      if (msg.indexOf("答题未通过") >= 0) {
        console.log("📢 需要答题，开始自动作答");
        if (doQuiz()) {
          sleep(1500);
          let r2 = clockIn();
          let h2 = handleResult(r2);
          if (h2 != null && h2.indexOf("🎉") >= 0) return "🎉 打卡成功\n";
          return h2 != null ? h2 : "❌ 打卡失败：" + (r2["msg"] || "未知错误") + "\n";
        }
        return "❌ 答题失败，无法完成打卡\n";
      }
      if (msg.indexOf("Captcha") >= 0 || msg.indexOf("ClientCode") >= 0 || msg.indexOf("Required") >= 0) {
        // 极端兜底：即便 ?member=wps 仍被要求微信验证码（需 wx.login()），优雅跳过
        return "⚠️ 渠道2打卡端点现需微信验证码(captcha/clientCode)，自动化无法完成；\n   请改用 WPS 微信小程序手动打卡，或在分表将该行「列F」设为「否」关闭本渠道\n";
      }
      return "❌ 打卡失败：" + msg + "\n";
    }

    // 先尝试答题（今日有题则作答，无题则跳过），再打卡；最大化兼容不同返回结构
    doQuiz();
    sleep(1200);
    let finalMsg = tryClockIn();
    if (finalMsg.indexOf("🎉") >= 0) {
      messageSuccess += finalMsg;
    } else {
      messageFail += finalMsg;
    }

  } catch {
    messageFail += "❌ " + messageName + "失败\n";
  }

  sleep(2000);
  if (messageOnlyError == 1) {
    messageArray[posLabel] = messageFail;
  } else {
    if(messageFail != ""){
      messageArray[posLabel] = messageFail + " " + messageSuccess;
    }else{
      messageArray[posLabel] = messageSuccess;
    }
  }

  if(messageArray[posLabel] != "")
  {
    console.log(messageArray[posLabel]);
  }
}
