//カレンダー作成スクリプト
//
//使い方
//node iCalCreate.js csvファイル.csv [0/1] [iCal/notion]
//
//csvファイル.csv 
//変換するファイル。スクリプト株のフォーマットに従う事。またDATELIST.csvも存在している必要がある。
//
//[0/1]
//0だと全体カレンダー用途(講師とカリキュラムを秘匿する)
//1だと個別カレンダー用途(講師とカリキュラムを表示する)
//省略可能で、デフォルトは0になる
//
//[iCal/notioncsv]
//iCalはカレンダーにインポート可能な形式で出力する
//notionはnotionにインポートする
//省略可能で、デフォルトはiCalになる

//////////////////////////////////////////////////////////////////////////
//
const fs = require('fs');
const iconv = require('iconv-lite');
const {parse} = require('csv-parse/sync');
const notion = require('./notion.js');
const { setTimeout } = require('timers/promises');


function getiCalTime(date)
{
  //ISO8601にする
  let ret = date.toISOString();
  ret = ret.replaceAll("-","");
  ret = ret.replaceAll(":","");
  ret = ret.replace(/\.(.*)Z/,"Z");
  return ret;
}

function mergeDate(day, time)
{
  let ret = new Date(day);
  let times = time.split(":");
  
  //ずれるのでUTCでは設定しない
  ret.setHours(times[0]);
  ret.setMinutes(times[1]);
  return ret;
}

const className = process.argv[2].replace(".csv","");
const classCsvText = fs.readFileSync(process.argv[2]);
const classCsvDecode = iconv.decode(classCsvText, "SJIS");
const classCsv = parse(classCsvDecode);
const dateCsv = parse(fs.readFileSync("DATELIST.csv"));

let title = "";      //カレンダータイトル
var output = "";
let outTitle = "";

let calType = 0;
let exportType = 0;
if(process.argv[3] == 1) {
  calType = 1;
}
if(process.argv[4] == "notion") {
  exportType = 1;
}

if(calType == 0) {
  title = "授業予定";
}
if(calType == 1) {
  title = className + "クラス授業予定";
}

if(exportType == 0)
{
  outTitle = process.argv[2].replace(".csv",".ical");
  output += "BEGIN:VCALENDAR\n";                //カレンダー始まり(変えない事)
  output += "PRODID: vtn-cal by k-mitarai\n";   //作成者
  output += "VERSION:2.0\n";                    //バージョン
  output += "CALSCALE:GREGORIAN\n";             //暦(グレゴリオ暦)
  output += `X-WR-CALNAME:${title}\n`;          //カレンダータイトル(Googleカレンダー用)
  output += "X-WR-TIMEZONE:Asia/Tokyo\n";       //タイムゾーン(Googleカレンダー用)
}

if(exportType == 1)
{
  outTitle = process.argv[2].replace(".csv","_importnotion.csv");
  
  //ヘッダ
  output += "開始時間,終了時間,場所,内容\n";
}

function createEventCal(evt, day)
{
  let startTime = mergeDate(day, evt.start);
  let endTime = mergeDate(day, evt.end);
  
  let st = getiCalTime(startTime);
  let ed = getiCalTime(endTime);
  let ts = getiCalTime(new Date());
  let summary = evt.summary;

  output += "BEGIN:VEVENT\n";             //予定はじまり(変えない事)
  output += `DTSTART:${st}\n`;            //開始時間
  output += `DTEND:${ed}\n`;              //終了時間
  output += `DTSTAMP:${ts}\n`;            //タイムスタンプ
  output += `LOCATION:${evt.location}\n`; //場所
  output += "SEQUENCE:0\n";               //修正回数
  output += "STATUS:CONFIRMED\n";         //確定済み予定
  output += `SUMMARY:${summary}\n`;       //内容
  output += "TRANSP:OPAQUE\n";            //透過度
  output += "END:VEVENT\n";               //予定終端(変えない事)

  //UID:Ical05de0472c73e6ee582f49b519fb0168e
  //CREATED:19000101T120000Z
  //LAST-MODIFIED:20230320T155912Z
}

async function createNotion(evt, day)
{
  let startTime = mergeDate(day, evt.start);
  let endTime = mergeDate(day, evt.end);
  let summary = evt.summary;
  
  let pageData = {
    "parent": {
      "database_id": "fdd978f86b1241e684fd9aa756131b98"
    },
    "properties": {
      "授業": {
        "title": [
          {
            "text": {
              "content": summary
            }
          }
        ]
      },
      "クラス": {
        "select": {
          "name" : className
        }
      },
      "場所": {
        "select": {
          "name" : evt.location
        }
      },
      "日時": {
        "date": {
          "start": startTime.toISOString(),
          "end": endTime.toISOString()
        }
      }
    }
  };
  
  await notion.createPage(pageData);
}

//作業用配列
let eventList = [];
for(let i=0; i<6; ++i){ eventList.push([]); }

//クラスの授業リストをイベントに分解
let week = 0;
let parent = 0;
let event = { name: "" };
let lastTimes = "";
for (const record of classCsv) {
  if(event.name != record[2])
  {
    let stTimes = record[1].split("-");
    
    if(event.name != "")
    {
      event.end = lastTimes[1];
      eventList[week].push(event);
    }
    
    if(record[2] != "")
    {
      event = {
        week: week,
        name: record[2],
        teacher: record[4],
        location: record[5],
        ext: record[6],
        start: stTimes[0],
      };
    }
    else
    {
      event = { name: "" };
    }
  }
  
  let time = parseInt(record[0]);
  if(parent > time)
  {
    week++;
  }
  parent = time;
  
  lastTimes = record[1].split("-");
}

(async() => {
//日程ごとにカレンダーを作る
week=0;
let interval={};
for (const record of dateCsv)
{
  if(record[0] == "") break;
  if(week > 5) week = 0;
  let time = parseInt(record[0]);
  
  if(eventList[week].length == 0)
  {
    week++;
    continue;
  }
  
  for(let e of eventList[week])
  {
    //隔週授業
    if(e.ext == "2")
    {
      if(interval[e.name])
      {
        interval[e.name] = 0;
        continue;
      }
      else
      {
        interval[e.name] = 1;
      }
    }
    
    
    let summary = "";
    if(calType == 0) {
      summary = className;
    }
    if(calType == 1) {
      summary = e.name;
    }
    
    //選択授業
    if(e.ext == "1" || e.ext == "2")
    {
      summary = "[選択授業]" + summary;
    }
    e.summary = summary;
    
    if(exportType == 0)
    {
      createEventCal(e, record[1]);
    }
    if(exportType == 1)
    {
      await setTimeout(200);
      await createNotion(e, record[1]);
    }
  }
  week++;
}

if(exportType == 0)
{
  output += "END:VCALENDAR\n";                  //カレンダーおわり(変えない事)
  fs.writeFileSync(outTitle,output);
}
})();


//授業日程csv
//以下の形式でもらう。
//何回目,日付
/*
1,2023/4/17
1,2023/4/18
1,2023/4/19
1,2023/4/20
1,2023/4/21
1,2023/4/22
2,2023/4/24
2,2023/4/25
2,2023/4/26
2,2023/4/27
2,2023/4/28
2,2023/5/13
*/


//クラス別csv
//月～土の順で並んでいる
//以下の形式でcsvをもらう。
//コマ,時間帯,授業名,クラス,講師,授業場所
/*
1,9:30-10:20,,,,
2,10:30-11:20,,,,
3,11:30-12:20,,,,
4,12:30-13:20,,,,
5,13:30-14:20,,,,
6,14:30-15:20,,,,
7,15:30-16:20,,,,
8,16:30-17:20,,,,
9,17:30-18:20,,,,
10,18:30-19:20,,,,
11,19:30-20:20,,,,
*/