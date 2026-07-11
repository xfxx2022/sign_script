/*
    name: "WPS轻量版签到"
    cron: 30 0 9 * * *
    脚本兼容: 金山文档（1.0），金山文档（2.0）
    更新时间：20260711
    环境变量名：无
    环境变量值：无
    备注：WPS轻量版/手机端签到，接口 vip.wps.cn/sign/v2。
          ✅ 已完善：签到前先调 get_data 预检查 is_sign，已签直接结束（不再误跑验证码重试）；
             10003 在预检查 is_sign=false 时才走验证码重试；连续 TryLater 判定为限流提前退出。
          配合“wps”分配置表使用（cookie 即 wps_sid）。
          “wps”分配置表列：A=cookie，B=是否执行，C=账号名称，
          D=转存PPT(稻壳版用)，E=是否渠道1打卡(本脚本用)，F=是否渠道2打卡，G=Signature(渠道2)
          说明：本脚本读取列E（是否渠道1打卡），为空或“是”时执行，填“否”则跳过。
*/

var sheetNameSubConfig = "wps"; // 分配置表名称（共享的分表，subConfigWps 的键）
let sheetNameSubConfig2 = "wps_light"; // 在 CONFIG 表中的唯一名称，用于读取个性化配置与写入消息队列
var pushHeader = "【WPS轻量版】"; // 推送头，给自己看的
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
    // 是否启用本渠道（渠道1 = 轻量版）。列E为空或“是”则执行，“否”则跳过。
    let channel1 = Application.Range("E" + pos).Text;
    if(channel1 != "" && channel1 != "是"){
      console.log("📢 " + messageName + " 未启用渠道1打卡，跳过");
      messageArray[posLabel] = "";
      return;
    }

    let url = "https://vip.wps.cn/sign/v2";
    let captchaUrl = "https://vip.wps.cn/checkcode/signin/captcha.png?platform=8&encode=0&img_witdh=275.164&img_height=69.184";
    headers = {
      "Cookie": "wps_sid=" + cookie,
      "Content-Type": "application/x-www-form-urlencoded",
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586"
    };

    // 解析响应，返回 { ok, result, msg, data }
    function parseSignResp(resp) {
      if (resp.status != 200) return { ok: false, result: "", msg: "HTTP " + resp.status, data: null };
      try {
        let j = resp.json();
        return { ok: true, result: j["result"] || "", msg: j["msg"] || "", data: j["data"] || null };
      } catch {
        return { ok: true, result: "", msg: "响应解析错误", data: null };
      }
    }

    // 预检查：查询今日是否已签（最可靠，避免把“今日已签”误判为失败）
    // 接口返回 data.is_sign 为 true 即今日已签到；该接口为社区 checkinpanel 标准做法
    function getSignStatus() {
      let sUrl = "https://vip.wps.cn/sign/mobile/v3/get_data";
      let sHeaders = {
        "Cookie": "wps_sid=" + cookie,
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586"
      };
      try {
        let resp = HTTP.get(sUrl, { headers: sHeaders });
        if (resp.status != 200) return { ok: false, reason: "HTTP " + resp.status };
        let j = resp.json();
        if (j && j.data && typeof j.data.is_sign !== "undefined") {
          return { ok: true, isSign: (j.data.is_sign === true || j.data.is_sign === 1) };
        }
        return { ok: false, reason: "无is_sign字段" };
      } catch (e) {
        return { ok: false, reason: "请求异常" };
      }
    }

    // 把服务端返回 data 里的奖励翻译成可读文案
    function formatReward(data) {
      if (!data) return "";
      let gift = data["gift_type"] || "";
      if (gift.indexOf("space_") == 0) {
        let mb = gift.substring(6);
        let extra = data["double"] == 1 ? "（已加倍）" : "";
        return mb + " MB 云空间" + extra;
      }
      return "奖励(" + JSON.stringify(data) + ")";
    }

    function reportResult(r) {
      if (r.result == "ok" || r.msg == "ok") {
        let reward = formatReward(r.data);
        messageSuccess += "🎉 签到成功" + (reward != "" ? "，获得 " + reward : "") + "\n";
        return true;
      }
      if (r.msg.indexOf("已经") >= 0 || r.msg.indexOf("已签") >= 0 || r.msg.indexOf("had") >= 0) {
        messageSuccess += "📢 今日已签到\n";
        return true;
      }
      return false;
    }

    // 把服务端错误码翻译成人话，避免只显示“10003”这种天书
    function describeFail(r) {
      let m = (r.msg == null ? "" : r.msg).toString();
      // 10003 在 WPS 轻量版单凭数字码无法区分，常见三种：①今日已签到 ②wps_sid 过期 ③风控验证码，给出自查指引
      if (m == "10003" || m == "") {
        return "未签到状态收到 10003：多为 wps_sid 过期需重新登录，或触发验证码风控。请去 vip.wps.cn 自查登录态，或稍后重试/手动签（已排除今日已签，因预检查 is_sign=false）";
      }
      if (m.indexOf("captcha") >= 0 || m.indexOf("验证码") >= 0) {
        return "被风控要求验证码，暂时无法自动通过，可关闭本账号列E或手动签到";
      }
      if (m.indexOf("已经") >= 0 || m.indexOf("已签") >= 0 || m.indexOf("had") >= 0) {
        return "今日已签到";
      }
      return "签到失败（服务端：" + m + "）";
    }

    // 0. 预检查今日是否已签（最可靠，避免把“今日已签”误判为失败）
    //    社区 checkinpanel 标准做法：先 GET get_data 看 is_sign，已签就直接结束，不浪费请求。
    let pre = getSignStatus();
    if (pre.ok && pre.isSign) {
      console.log("📢 预检查：今日已签到，直接结束");
      messageSuccess += "📢 今日已签到\n";
    } else {
      if (!pre.ok) {
        console.log("⚠️ 预检查不可用（" + (pre.reason || "") + "），继续走真实签到流程");
      }
      // 1. 先尝试不带验证码坐标（部分账号/时段可直接免验证）
      console.log("📡 轻量版：先尝试免验证签到");
      let resp0 = HTTP.post(url, { "platform": "8" }, { headers: headers });
      let r0 = parseSignResp(resp0);
      console.log(r0);
      if (reportResult(r0)) {
        // 免验证直接成功或今日已签到，结束
      } else if (r0.result == "error" && (r0.msg == "10003" || r0.msg == "")) {
        // 2. 免验证未通过（10003/空msg）：可能是需验证码风控或 wps_sid 失效。
        //    注意：此处已排除“今日已签”（预检查 is_sign=false），所以 10003 是真失败，走验证码重试。
        console.log("📡 免验证未通过（" + (r0.msg || "") + "），尝试刷新验证码并带坐标重试");
        let dataWithCaptcha = {
          "platform": "8",
          "captcha_pos": "137.00431974731889, 36.00431593261568",
          "img_witdh": "275.164",
          "img_height": "69.184"
        };
        let ok = false;
        let rateLimited = false;
        for (let n = 0; n < 10; n++) {
          try {
            HTTP.get(captchaUrl, { headers: headers }); // 触发服务端刷新验证码
          } catch {}
          sleep(300);
          let resp = HTTP.post(url, dataWithCaptcha, { headers: headers });
          let r = parseSignResp(resp);
          console.log("第" + (n + 1) + "次带坐标尝试：" + JSON.stringify(r));
          if (reportResult(r)) {
            ok = true;
            break;
          }
          if (r.msg == "TryLater") {
            // 被限流：同一 sid 短时间内狂刷验证码接口会触发，继续重试无意义，提前退出
            rateLimited = true;
            break;
          }
          sleep(300);
        }
        if (!ok) {
          if (rateLimited) {
            messageFail += "❌ " + describeFail(r0) + "（验证码重试被限流 TryLater，请稍后或手动签）\n";
          } else {
            messageFail += "❌ " + describeFail(r0) + "（带坐标重试 10 次未通过）\n";
          }
        }
      } else {
        // 其他错误（如 captcha/验证码文案、未知码）
        messageFail += "❌ " + describeFail(r0) + "\n";
      }
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
