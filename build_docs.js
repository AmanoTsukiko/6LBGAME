const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  Header, Footer, LevelFormat, PageNumber, HeightRule,
} = require("C:/Users/chleung/3LA3-work/node_modules/docx");

const OUT = __dirname;
const KAI = "標楷體";
const PAGE_W = 11906;
const PAGE_H = 16838;
const MARGIN = 720;
const CW = PAGE_W - MARGIN * 2;

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const hair = { style: BorderStyle.SINGLE, size: 4, color: "666666" };
const box = { top: hair, bottom: hair, left: hair, right: hair };
const noB = { top: none, bottom: none, left: none, right: none };

function kai(text, size, extra) {
  return new TextRun(Object.assign({
    text: text,
    font: { ascii: KAI, eastAsia: KAI, hAnsi: KAI },
    size: size,
  }, extra || {}));
}

function p(text, size, opts) {
  const o = opts || {};
  const children = Array.isArray(text) ? text : [kai(text, size, o.run)];
  return new Paragraph({
    alignment: o.align || AlignmentType.LEFT,
    spacing: { before: o.before == null ? 60 : o.before, after: o.after == null ? 60 : o.after, line: o.line || 360, lineRule: "auto" },
    border: o.border,
    pageBreakBefore: !!o.breakBefore,
    keepNext: !!o.keepNext,
    keepLines: !!o.keepLines,
    numbering: o.numbering,
    children: children,
  });
}

function runs(parts, size) {
  return parts.map(function (part) {
    if (typeof part === "string") return kai(part, size);
    return kai(part.text, part.size || size, part);
  });
}

function blankLine() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 1 } },
    spacing: { before: 40, after: 40, line: 360, lineRule: "auto" },
    children: [kai(" ", 32)],
  });
}

function cell(text, width, opts) {
  const o = opts || {};
  const size = o.size || 32;
  const paras = Array.isArray(text) ? text : [p(text, size, { before: 40, after: 40, align: o.align })];
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: box,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    shading: { type: ShadingType.CLEAR, fill: o.fill || "FFFFFF" },
    verticalAlign: o.valign || VerticalAlign.TOP,
    children: paras,
  });
}

function plainCell(paragraphs, width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: noB,
    margins: { top: 20, bottom: 20, left: 60, right: 60 },
    children: paragraphs,
  });
}

function schoolHeader(objectiveLines, sheetTitle) {
  const out = [
    p("佛教黃焯菴小學", 32, { align: AlignmentType.CENTER, before: 0, after: 0, run: { bold: true } }),
    p("科學科工作紙", 32, { align: AlignmentType.CENTER, before: 0, after: 40, run: { bold: true } }),
  ];
  objectiveLines.forEach(function (line) {
    out.push(p(line, 32, { align: AlignmentType.CENTER, before: 0, after: 0 }));
  });
  const leftW = 5233;
  const rightW = CW - leftW;
  out.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [leftW, rightW],
    borders: {
      top: none, bottom: none, left: none, right: none,
      insideHorizontal: none, insideVertical: none,
    },
    rows: [
      new TableRow({
        children: [
          plainCell([p("六年級_____班", 32, { before: 60, after: 20 })], leftW),
          plainCell([p("成績：_______________", 32, { before: 60, after: 20, align: AlignmentType.RIGHT })], rightW),
        ],
      }),
      new TableRow({
        children: [
          plainCell([p("姓名：_______________ (  )", 32, { before: 20, after: 40 })], leftW),
          plainCell([p("日期：_______________", 32, { before: 20, after: 40, align: AlignmentType.RIGHT })], rightW),
        ],
      }),
    ],
  }));
  out.push(p(sheetTitle, 32, {
    align: AlignmentType.CENTER, before: 80, after: 80, run: { bold: true },
  }));
  return out;
}

function dataTable(columnWidths, rows) {
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: columnWidths,
    rows: rows.map(function (row) {
      return new TableRow({
        cantSplit: row.cantSplit !== false,
        height: row.height ? { value: row.height, rule: HeightRule.ATLEAST } : undefined,
        tableHeader: !!row.header,
        children: row.cells.map(function (item, index) {
          const cellItem = (item && typeof item === "object" && !Array.isArray(item) && item.text != null) ? item : { text: item };
          const text = cellItem.text;
          const fill = cellItem.fill || (row.header ? "E7EFE8" : "FFFFFF");
          const paras = Array.isArray(text) ? text : [p(String(text), cellItem.size || 32, {
            before: 40, after: 40,
            align: cellItem.align || (row.header ? AlignmentType.CENTER : AlignmentType.LEFT),
            run: (cellItem.bold || row.header) ? { bold: true } : undefined,
          })];
          return cell(paras, columnWidths[index], {
            fill: fill,
            valign: cellItem.valign || VerticalAlign.CENTER,
          });
        }),
      });
    }),
  });
}

function lines(items, size) {
  return items.map(function (item, index) {
    return p(item, size || 28, { before: index === 0 ? 40 : 20, after: 20 });
  });
}

function page() {
  return {
    size: { width: PAGE_W, height: PAGE_H },
    margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
  };
}

const VIDEOS = [
  {
    animal: "海龜",
    title: "Sea Turtles 101 | National Geographic",
    url: "https://www.youtube.com/watch?v=5Rmv3nliwCs",
    time: "3分45秒",
    watch: "全片。認識海龜，並記下片中人類怎樣令海龜變少。",
  },
  {
    animal: "海龜（漁網與沙灘，時間不足可改播這條）",
    title: "Protecting marine turtles in Coastal Kenya | WWF UK",
    url: "https://www.youtube.com/watch?v=_xv9p66beA0",
    time: "3分49秒",
    watch: "全片。漁網、產卵沙灘、人們怎樣對待海龜。",
  },
  {
    animal: "老虎",
    title: "What's Driving Tigers Toward Extinction? | National Geographic",
    url: "https://www.youtube.com/watch?v=qhrT0v7QSvc",
    time: "5分46秒；課堂播 0:00–2:40",
    watch: "捕獵和虎製品。2:40 可停。",
  },
  {
    animal: "大象",
    title: "Why Elephants May Go Extinct in Your Lifetime | National Geographic",
    url: "https://www.youtube.com/watch?v=zJbFthhZRHY",
    time: "1分52秒",
    watch: "全片。象牙和捕獵。",
  },
];

function videoTable() {
  const cols = [1400, 4200, 4866];
  const head = new TableRow({
    tableHeader: true,
    children: ["動物", "片名與網址", "怎樣看"].map(function (label, i) {
      return cell(label, cols[i], { fill: "E7EFE8", size: 24, align: AlignmentType.CENTER });
    }),
  });
  const rows = VIDEOS.map(function (video) {
    return new TableRow({
      children: [
        cell(video.animal, cols[0], { size: 22 }),
        cell([
          p(video.title, 22, { before: 20, after: 20, run: { bold: true } }),
          p(video.url, 20, { before: 0, after: 20 }),
          p(video.time, 20, { before: 0, after: 20 }),
        ], cols[1]),
        cell(video.watch, cols[2], { size: 22 }),
      ],
    });
  });
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: cols,
    rows: [head].concat(rows),
  });
}

function h(text, opts) {
  return p(text, 32, Object.assign({ before: 200, after: 80, run: { bold: true } }, opts || {}));
}

function body(text, opts) {
  return p(text, 24, Object.assign({ before: 40, after: 40, line: 360 }, opts || {}));
}

function teacherDoc() {
  const children = [];
  children.push(p("佛教黃焯菴小學", 36, { align: AlignmentType.CENTER, before: 0, after: 40, run: { bold: true } }));
  children.push(p("科學科　六年級　教師課節流程", 32, { align: AlignmentType.CENTER, before: 0, after: 40, run: { bold: true } }));
  children.push(p("瀕危物種", 36, { align: AlignmentType.CENTER, before: 80, after: 40, run: { bold: true } }));
  children.push(p("能力較高班與能力稍遜班", 28, { align: AlignmentType.CENTER, before: 0, after: 120 }));

  children.push(h("本科目標"));
  children.push(body("6LD1 認識一些瀕危物種"));
  children.push(body("6LD2 解釋瀕危物種瀕臨絕種的原因"));

  children.push(h("給同事：這一課的順序"));
  children.push(body("兩班都依同一順序走完，不要調換。順序是 Do-it，然後 Talk-it，然後 Read-it，然後 Write-it。"));
  children.push(body("Do-it 讓學生先遇見「數目會下降」這件事。Talk-it 讓學生在打開課文之前寫下主張。Read-it 才讀三篇原文，用 KWL 把已知、想知和讀到的分開。Write-it 回到自己的主張，看文本的證據支持還是不支持。"));
  children.push(body("建議用四課，每課走一步。第一課 Do-it，用動手記錄。第二課 Talk-it，在主張紙寫主張，主張紙留到第四課。第三課 Read-it，主張寫好之後才發 KWL。第四課 Write-it，才發寫作工作紙，把主張原句抄回去。若只有三課，第四課接在第三課讀完之後。"));

  children.push(h("這一課用的閱讀方法"));
  children.push(body("Seeds of Science/Roots of Reading（Lawrence Hall of Science，加州大學柏克萊分校；Cervetti, Pearson, Barber 與同事）。科學和閱讀織在一起，四個動作是 Do-it、Talk-it、Read-it、Write-it。第一手探究 firsthand inquiry 是親身觀察、實驗、模擬和搜集資料。第二手探究 secondhand inquiry 是閱讀並判斷文本。兩者沒有固定先後，這一課按本科需要採用 Do-it、Talk-it、Read-it、Write-it。"));
  children.push(body("捕獵遊戲是模擬的 firsthand Do-it：學生當獵人，對準動物射擊，射中才少一隻。能力較高班再以影片表做 Do-it。能力稍遜班再以角色扮演做 Do-it。三篇原文是 secondhand Read-it。"));
  children.push(body("KWL（Ogle, 1986：Know, Want to know, Learned）。讀文前填 K「我知道什麼」和 W「我想知道什麼」。讀文後填 L「我讀到什麼」。K 可以包括遊戲、影片或角色扮演裏已經知道的，也包括未必正確的想法，先留著。W 要寫問題，不寫課題名稱。L 只寫文本裏讀到的。"));
  children.push(body("動手記錄用表格記下第一手觀察：開始有多少隻、射中後還剩多少、沒有射中時數目有沒有變。先記觀察，再談原因。KWL 是三欄表，每欄寫明何時填。寫作工作紙是主張、證據、解釋三欄。主張在讀文前寫好，讀文後才用課文判斷。"));

  children.push(h("課前放在同一資料夾"));
  children.push(body("hunting-game.html　用瀏覽器直接打開，不用上網，也不用安裝。"));
  children.push(body("學生工作紙頁首依教學指引附件二：校名、科學科工作紙、兩個目標、班別與成績、姓名與日期，下面才是工作紙名稱。"));
  children.push(body("6LD_動手記錄.docx　第一課，兩班都發。兩班都填捕獵遊戲表。能力較高班再填影片表。能力稍遜班再填角色扮演表。"));
  children.push(body("6LD_主張紙.docx　第二課，打開閱讀之前才發。這時不要發 KWL，也不要發寫作工作紙。若當天沒有印主張紙，可先寫在便利貼，仍然不要發後面兩份。"));
  children.push(body("6LD_KWL科學閱讀.docx　第三課、主張已經寫好之後才發。這份閱讀的後面就是三篇原文，提早發出，學生會先看見原因。能力較高班按組派一種動物；能力稍遜班三種都發，由老師帶讀。"));
  children.push(body("6LD_寫作工作紙.docx　第四課才發。學生把主張紙上的原句抄到主張表，再寫證據和解釋。"));
  children.push(body("投影片：6LD_能力較高班_課件.pptx、6LD_能力稍遜班_課件.pptx。跟著投影片上的「老師做」和「學生做」走即可。"));

  children.push(h("能力較高班　Do-it　影片"));
  children.push(body("影片是英文。老師用廣東話在旁說明，學生不用抄英文字幕。每條片播完先暫停，讓學生在動手記錄的影片表寫三樣：我看見或聽見的一件事、人類做了什麼令牠變少、我還想知道的一個問題。"));
  children.push(videoTable());
  children.push(body("播放順序：海龜 Sea Turtles 101，老虎 What's Driving Tigers Toward Extinction? 的開首 2 分 40 秒，大象 Why Elephants May Go Extinct in Your Lifetime。若海龜那條的絕種部分聽不清楚，改播 WWF UK 的 Protecting marine turtles in Coastal Kenya，內容是漁網和產卵沙灘。"));
  children.push(body("老虎原文第 7 題指向片中英文 “When the buying stops, the killing can too.”。今天這幾條片沒有這一句。不要為了這一句另找片段。讀文時告訴學生這句不在今天的影片裏。"));

  children.push(h("能力較高班逐段", { breakBefore: true }));
  children.push(body("第一課 Do-it，約一教節。先打開 hunting-game.html。讓全班看見海龜 8 隻、老虎 6 隻、大象 6 隻。學生是獵人：對準畫面上的動物，按一下滑鼠，射中才少一隻。沒有射中，數目不變。每射中一次，停下數還剩多少。學生在動手記錄寫下開始的隻數、射中後還剩多少、沒有射中時數目有沒有變。問：數目怎樣變？先聽學生說。然後播上面三條片。看完一條，先在影片表寫，再播下一條。老虎片播到 2 分 40 秒。"));
  children.push(body("老師可以說：我們先當獵人，看發生什麼事。射中才少一隻，沒有射中數目不變。每射中一次，數一數還剩多少。等一下再說原因。"));
  children.push(body("第二課 Talk-it，約一教節。先用幾分鐘回想遊戲和影片裏看見的事，再討論動物瀕臨絕種的可能原因。每人在主張紙寫多於一個主張，至少兩個，兩個要是不同原因。句式：我主張（動物）瀕臨絕種，是因為（原因）。這時不要發 KWL，也不要發寫作工作紙。打開課文之前不要公布原文的原因清單，也不要塗掉主張。主張紙收好，第四課發還。"));
  children.push(body("第三課 Read-it，約一教節。主張已經寫好，這時才發 KWL 科學閱讀。讀文前先填 K 和 W。分組：每一組只讀一種動物，該動物的全部版本都要讀。可以有多於一組讀同一種動物。老師分配時盡量三種都有組別，匯報時全班才聽得見三種原因。讀的時候畫出人類令動物變少的句子。讀完填 L，L 只寫自己那篇。證據表先填自己的動物，另外兩種聽匯報後各抄一句，並圈和自己的主張有關或無關。每組用一句話匯報一個人為原因，並說出它出現在文本的哪一處。"));
  children.push(body("第四課 Write-it，約一教節。這時才發寫作工作紙。把主張紙上的主張原句抄到主張表，不要改成課文。物種表：自己讀的那種寫一項特徵和一個人為原因，另外兩種寫匯報聽到的。然後做主張、證據、解釋。文本支持就抄證據，並寫這句證據怎樣支持主張。文本說的是另一個原因，就圈「文本不支持」，寫出文本實際的原因。最後用檢查表看三樣是否都寫了。"));

  children.push(h("能力稍遜班逐段"));
  children.push(body("第一課 Do-it，約一教節。全班一起看 hunting-game.html，由老師或一位學生當獵人，對準動物按一下滑鼠。每射中一次，全班齊聲說還剩多少。沒有射中，數目不變。在動手記錄寫下射中後還剩多少、沒有射中時數目有沒有變。然後做角色扮演。三種動物都要有人演。每人只演一種。動手記錄和投影片上都有台詞，學生跟著說，並加上動作。演完在「我還想到的原因」寫一個自己想到的原因。"));
  children.push(body("海龜：雙手作划水。說：我回到沙灘產卵。魚網困住我。人們拿走我的蛋，又把沙灘建成旅遊區。"));
  children.push(body("老虎：雙手作爪。說：森林被砍掉，我沒有地方住。人們捕獵我，取去皮和骨頭。"));
  children.push(body("大象：一隻手臂作象鼻。說：人們佔去森林。他們捕獵我，只為了象牙。"));
  children.push(body("角色扮演用的是這三句台詞，方便學生開口。判斷對錯時仍以三篇原文為準，不把台詞當成課文。"));
  children.push(body("第二課 Talk-it，約一教節。先回想遊戲和角色扮演，再討論原因。每人只在主張紙寫一個主張。句式：我主張這些動物瀕臨絕種，是因為______。可以寫角色扮演或遊戲裏聽過的一個原因。主張二留空。這時不要發 KWL，也不要發寫作工作紙。主張紙收好，第四課發還。"));
  children.push(body("第三課 Read-it，約一教節。主張已經寫好，這時才發 KWL 科學閱讀。不按動物分組。老師帶領全班讀三種動物。每種先讀簡介裏的一兩句特徵，再讀「面臨的威脅」。讀出原因句時讓學生用筆畫線。打開文本前填 K 和 W；每讀完一種，在 L 寫一個原因。證據表三種都填，每種抄一句人類令牠變少的句子，並圈和自己的主張有關或無關。"));
  children.push(body("帶讀時指住這些句子。海龜：意外地被綿長的捕魚網所困；生吃龜肉；偷吃牠們的蛋；產卵的沙灘發展為旅遊區；龜皮製成靴及手袋；標本；殼上的物質製成首飾。老虎：破壞森林、發展農業、斬砍木材，失去棲息之所；皮毛；虎骨；非法偷獵。大象：毀林造屋，棲息地縮小；捕獵大象以滿足象牙市場。"));
  children.push(body("第四課 Write-it，約一教節。這時才發寫作工作紙。把主張紙上的那一句原句抄到主張一，主張二留空。物種表三種各寫一項特徵和一個原因。主張、證據、解釋只做主張一。圈文本支持或不支持，從閱讀抄一句證據，並寫這句證據怎樣和主張有關。若文本不支持，寫出文本實際的原因。"));

  children.push(h("主張、證據和解釋"));
  children.push(body("寫作工作紙的三欄是主張、證據、解釋，對應 Claim, Evidence, Reasoning（McNeill & Krajcik, 2012）。主張回答「這些動物為什麼瀕臨絕種」，而且在讀文之前已經寫在主張紙上。證據是從課文抄來的一句，並寫上動物名稱。遊戲、影片和角色扮演用來想起主張；判斷對錯仍以三篇原文為準。解釋用「這句證據支持（或不支持）我的主張，因為……」把證據和主張接起來。"));

  children.push(h("怎樣看學生的主張"));
  children.push(body("「因為被人捕獵」三種文本都有證據。海龜是漁網、龜肉、龜蛋、龜皮、標本和龜殼首飾。老虎是皮毛、虎骨和非法偷獵。大象是象牙。"));
  children.push(body("「因為棲息地減少」老虎和大象的文本支持。海龜文本寫的是漁網、肉、蛋、沙灘和殼製品。只讀了海龜而主張棲息地減少，應判斷為文本不支持，並寫出海龜文本的原因。"));
  children.push(body("只寫「因為會絕種」或把主張改寫成課文原句，不能算完成。學生要保持讀文前的主張，再用文本判斷。"));

  children.push(h("原文照讀，數字不要換"));
  children.push(body("簡介已改成小六能讀的短句。重量、壽命、食物和「小海龜在兩日後破殼而出」等事實仍保留。面臨威脅的段落沒有改，仍是原文。老虎文用「本世紀初」、約十萬頭、餘下約五千頭。大象文寫 1970 年至 1989 年由 150 萬頭減至不足 60 萬頭，並寫亞洲象牙工業以香港為基地。這些數字可能過時，課堂上不要換成較新的數目。"));
  children.push(body("老虎（一）原文註明改編自 http://www.jackiewild.com ，問題由王志平及麥思源編寫。"));

  children.push(h("參考文獻"));
  children.push(body("Cervetti, G. N., Pearson, P. D., Barber, J., Hiebert, E. H., & Bravo, M. A. (2007). Integrating literacy and science: The research we have, the research we need. In M. Pressley, A. K. Billman, K. H. Perry, K. E. Reffitt, & J. M. Reynolds (Eds.), Shaping literacy achievement (pp. 157–174). New York: Guilford."));
  children.push(body("Lawrence Hall of Science, University of California, Berkeley. Seeds of Science/Roots of Reading. Do-it, Talk-it, Read-it, Write-it；firsthand inquiry 與 secondhand inquiry。"));
  children.push(body("Ogle, D. M. (1986). K-W-L: A teaching model that develops active reading of expository text. The Reading Teacher, 39(6), 564–570."));
  children.push(body("McNeill, K. L., & Krajcik, J. (2012). Supporting grade 5–8 students in constructing explanations in science: The claim, evidence, and reasoning framework for talk and writing. Boston: Pearson."));

  return new Document({
    styles: { default: { document: { run: { font: KAI, size: 24 } } } },
    sections: [{
      properties: { page: page() },
      children: children,
    }],
  });
}

const READINGS = [
  {
    animal: "海龜",
    versions: [
      {
        title: "海龜（一）簡介",
        paras: [
          "海龜可重達 600 公斤，身長 1.8 米，比其他爬蟲類重。牠們可以活 100 年以上。牠們吃海藻、水母、蟹和魚。牠們在各地的大海生活，由極北到極南，但只在暖水交配。",
          "海龜每年要游數千里，前往覓食的地方。海龜住在海裡，但仍要呼吸空氣。牠們用腳掌划水，游得很快。牠們可以把心跳減到每九分鐘一下，因此能連續游數小時，不用換氣。",
          "雌海龜每年回到同一個沙灘產卵，沒有人知道牠們怎樣認得路。海龜媽媽爬上沙灘，挖一個大洞，一次可產卵 100 枚。牠把卵埋在沙裡，然後回到海裡。小海龜在兩日後破殼而出，爬出沙堆，靠月光下海。海龜要到 50 歲才能交配。少於 1% 的海龜能活到成年。",
          "海龜是古老的爬蟲類，已有 1 億 5 千萬年歷史，比恐龍更早在地球出現。現在有 7 種海龜。牠們大多在海上活動，數目難以估計。可以肯定的是，7 種海龜都是瀕危動物。",
        ],
      },
      {
        title: "海龜（二）海龜面臨的威脅",
        paras: [
          "Leatherback 海龜是最大的海龜，成年雌龜不足 10 萬隻。Loggerhead 海龜長年住在中東和美國東南部，成年雌龜只餘 2.8 萬隻。Kemp's Ridley 海龜是最瀕危的一種。Flatback 海龜只見於澳洲水域，雌龜約有 5 至 10 萬隻，雄龜的數目不知道。",
          "海龜面對的困境，主要是由人類造成的。牠們會意外地被綿長的捕魚網所困。有些地區，人們喜歡生吃龜肉，亦會偷吃牠們的蛋；將牠們產卵的沙灘發展為旅遊區，破壞沙灘的舊貌，使牠們無法產卵。",
          "亦有為供應消費市場而殺害海龜的。例如，人們以龜皮製成靴及手袋；亦有把牠們製成標本當紀念品發售的，有的甚至把牠們製成樂器。更不幸的是海龜背部的殼上還有一層透明的美麗物質，可製成首飾。",
          "我們如不馬上行動起來，海龜便會於不久的將來自地球表面永久消失。我們必須停止一切有關海龜的買賣，及支持保護牠們生態環境的計劃。",
        ],
      },
    ],
  },
  {
    animal: "老虎",
    versions: [
      {
        title: "老虎（一）簡介",
        note: "文章改編自 http://www.jackiewild.com　問題由王志平及麥思源編寫。",
        paras: [
          "老虎是體型最大的貓科動物。由鼻子到尾巴長 3 米，體重可達 170 公斤。每頭老虎的條紋都不相同，可以靠條紋分辨牠們。條紋的顏色幫助牠們在草叢中藏身，方便捕獵。",
          "老虎吃肉，爪和牙齒很有力。牠們多在夜間獵食鹿、羚羊、野豬和其他小動物。牠們一餐可吃 30 公斤肉，然後可以一個星期不再進食。野生老虎大約活 15 年。牠們喜歡獨居，只在交配時才聚在一起。雌虎每次生下 2 至 3 隻幼虎。幼虎跟母親學習捕獵，到 18 個月才獨立生活。",
          "老虎在地球上已有二百萬年。從前牠們生活在亞洲不同的地方，由雪山、森林、草原到熱帶雨林和沿海沼澤都有。今天，牠們的活動範圍小了很多，只分散在一些小角落。",
        ],
      },
      {
        title: "老虎（二）老虎面臨的威脅",
        paras: [
          "直至本世紀初，仍有八個亞種約十萬頭老虎存在於世上。時至今天，其中三個亞種已經絕滅，老虎的數目亦銳減至僅餘五千頭。在餘下的五個亞種中，南中國虎不足 50 頭，西伯利亞虎不足 300 頭，蘇門答臘虎不足 400 頭。有為數約一千頭印老虎活躍於東南亞，而數目最多的野生老虎，生存於印度及尼泊爾一帶。",
          "由於人類破壞森林，發展農業，以解決人口日益膨脹的糧食問題，又大量斬砍木材用來製造各種用具，滿足生活需求。老虎因為失去棲息之所而日見稀少。",
          "從前西方人會謀取老虎的皮毛作為戰利品或造成地毯；而在東方，中國人認為老虎身體的每一部分，均有藥用價值。例如，虎骨可浸酒可治風濕。不過，並沒有確實的科學證據，證明虎骨的醫療功效。這種誤傳，很可能只是古時中國人以為老虎勇猛過人，吃其筋骨可得其力量。",
          "自 1976 年，華盛頓公約已禁止一切涉及老虎身體部份的貿易。雖然老虎是受保護的動物，不過牠們現時面臨的最大威脅，仍是來自非法的偷獵活動。",
        ],
      },
      {
        title: "老虎（三）老虎買賣",
        paras: [
          "於本世紀初估計存活於中國的老虎仍在四萬之數。今天，沒有人知道仍然留存於中國境內老虎的確實數量，只知碩果僅存的華南虎不足 50 頭。由於含有老虎成分的藥物價錢極為高昂，所以現今仍有不法商人遠赴印度、俄羅斯和東南亞尋找貨源，以滿足用家的需求。",
          "世界八成老虎生存於印度，而印度政府對保育老虎亦不遺餘力。官方於境內開闢了 100 個老虎自然保護區。可是，偷獵仍時有發生，而走私活動更層出不窮。估計於印度每天即有二頭老虎遭獵殺。雖然，印度的偷獵者打死一頭老虎，可能只得幾百美元的報酬，但對他而言，便是一年的收入了。",
          "當蘇聯解體時，有大批以售賣野生動物為生的人由中韓邊界湧入俄國，大量偷獵近邊界區域的西伯利亞虎，和進行有關的走私活動。當時俄國政府缺乏資源去加以保護，這亞種老虎瀕危的處境十分令人憂慮。於 1994 年，西伯利亞老虎的數目因而銳減至不足 150 頭。幸好於同年，俄羅斯成立了救虎行動組（amba），它是由四個反偷獵的武裝人員組成的。組員除巡查老虎的棲息地和禁止狩獵活動外，更於邊境阻截走私人士，以保護碩果僅存的老虎。據悉，該組已成功減少 75% 的非法捕獵。科學家估計，野生西伯利亞老虎已增加至大約 300 頭。",
        ],
      },
    ],
  },
  {
    animal: "大象",
    versions: [
      {
        title: "大象（一）",
        paras: [
          "大象可以活 60 年，體重可達 6 噸。象鼻是大象特有的。大象用象鼻喝水、洗澡、聞氣味、呼吸和找食物。象牙長在象鼻兩旁，和人的牙齒一樣。象牙可以用來撕開樹皮、搬重的東西和挖地找水。",
          "大象不吃肉。牠們吃草、水果、樹葉和樹枝。牠們每天要用約 16 小時，尋找 160 公斤食物和 50 公升水。",
          "大象群居。象群由最年長、最有經驗的雌象帶領，一起保護和照顧小象。",
          "大象已有數百萬年歷史。今天主要有非洲象和亞洲象兩種。",
          "非洲象是陸地上最大的動物，生活在非洲的森林和草原。現存的非洲象不足 60 萬頭。",
          "亞洲象體型較小，耳朵和象牙也較小，前額圓，背部隆起。牠們生活在印度、斯里蘭卡、中國、印度尼西亞和東南亞的森林裡。估計現存約 4 萬頭。長久以來，人們訓練亞洲象做粗重的工作，例如運送木材和其他貨物。",
        ],
      },
      {
        title: "大象（二）大象面臨的威脅",
        paras: [
          "人類不斷擴張，毀林造屋及侵入森林地區居住，造成大象的天然棲息地日漸縮小，影響其生存。可是，對非洲大象最大的威脅，卻來自偷獵者的彈頭。人們捕獵大象以滿足龐大的象牙市場。僅在 1970 年至 89 年間，非洲大象由原來的 150 萬頭，減少至不足 60 萬頭。在非洲地區發現的象群屍體，只有象牙被掠走。",
          "人類視象牙為珍貴的材料，可雕刻成裝飾品及首飾，亦用於鋼琴的琴鍵及製造印章。過去，亞洲的象牙工業以香港為基地，採用現代工具大量生產，單單一間工廠，便能於一週內處理過百枝象牙，以滿足亞洲市場的需求。",
          "華盛頓公約組織及國際社會於 1989 年全面禁止象牙貿易，偷獵的情況已經有所放緩，使美國及歐洲的象牙市場式微。但在亞洲，雕塑仍見於貨架上，有關象牙買賣仍在廣泛進行。如果亞洲人都拒絕購買象牙製品，其價值便不復從前，而大象便可望逃過新一輪的屠殺。",
        ],
      },
    ],
  },
];

const OBJECTIVES = [
  "6LD1 認識一些瀕危物種",
  "6LD2 解釋瀕危物種瀕臨絕種的原因",
];

function studentDoc(children) {
  return new Document({
    styles: { default: { document: { run: { font: KAI, size: 32 } } } },
    sections: [{ properties: { page: page() }, children: children }],
  });
}

function sheetHead(title) {
  return schoolHeader(OBJECTIVES, title);
}

function clipLabel(animal, title, url) {
  return [
    p(animal, 32, { before: 20, after: 8, run: { bold: true } }),
    p(title, 20, { before: 0, after: 0 }),
    p(url, 18, { before: 0, after: 12 }),
  ];
}

function room(prompt) {
  return [p(prompt, 32, { before: 40, after: 40 })];
}

function doDoc() {
  const children = [];
  children.push.apply(children, sheetHead("動手記錄"));
  children.push(dataTable([3488, 3489, 3489], [
    { header: true, cells: ["兩班都填", "看影片的再填", "做角色扮演的再填"] },
    { cells: ["下面的獵人模擬表", "影片表", "角色扮演表"] },
  ]));
  children.push(p("你是獵人。對準動物按一下。", 32, {
    before: 80, after: 0, keepNext: true, keepLines: true,
  }));
  children.push(p("射中了，寫還剩多少。沒有射中，寫數目有沒有變。", 32, {
    before: 0, after: 40, keepNext: true, keepLines: true,
  }));
  children.push(dataTable([1600, 1800, 3533, 3533], [
    { header: true, cells: ["動物", "開始（隻）", "射中後還剩（隻）", "沒有射中，有沒有變"] },
    { height: 1100, cells: ["海龜", "8", room(" "), room(" ")] },
    { height: 1100, cells: ["老虎", "6", room(" "), room(" ")] },
    { height: 1100, cells: ["大象", "6", room(" "), room(" ")] },
  ]));

  children.push(p("看影片的一班", 32, { before: 200, after: 40, run: { bold: true }, keepNext: true }));
  children.push(p("看完一條先寫，再播下一條。老虎片看到 2 分 40 秒。問題留到科學閱讀的 W。", 32, { before: 0, after: 60, keepNext: true, keepLines: true }));
  const clips = [
    ["海龜", "Sea Turtles 101 | National Geographic", "https://www.youtube.com/watch?v=5Rmv3nliwCs"],
    ["老虎", "What's Driving Tigers Toward Extinction? | National Geographic", "https://www.youtube.com/watch?v=qhrT0v7QSvc"],
    ["大象", "Why Elephants May Go Extinct in Your Lifetime | National Geographic", "https://www.youtube.com/watch?v=zJbFthhZRHY"],
  ];
  children.push(dataTable([3200, 2422, 2422, 2422], [
    { header: true, cells: ["動物和片名", "我看見或聽見", "人類做了什麼", "我還想知道"] },
  ].concat(clips.map(function (clip) {
    return {
      height: 1700,
      cells: [
        clipLabel(clip[0], clip[1], clip[2]),
        room(" "),
        room(" "),
        room(" "),
      ],
    };
  }))));
  children.push(p("若海龜片的威脅聽不清楚，改看 Protecting marine turtles in Coastal Kenya | WWF UK", 28, { before: 80, after: 10 }));
  children.push(p("https://www.youtube.com/watch?v=_xv9p66beA0", 24, { before: 0, after: 40 }));

  children.push(p("做角色扮演的一班", 32, { breakBefore: true, before: 0, after: 40, run: { bold: true } }));
  children.push(p("每人演一種。跟著台詞說，再加上動作。演完寫一個自己想到的原因。", 32, { before: 0, after: 60 }));
  const roles = [
    ["海龜", "雙手作划水", "我回到沙灘產卵。魚網困住我。人們拿走我的蛋，又把沙灘建成旅遊區。"],
    ["老虎", "雙手作爪", "森林被砍掉，我沒有地方住。人們捕獵我，取去皮和骨頭。"],
    ["大象", "一隻手臂作象鼻", "人們佔去森林。他們捕獵我，只為了象牙。"],
  ];
  children.push(dataTable([1400, 2000, 4266, 2800], [
    { header: true, cells: ["動物", "動作", "台詞", "我還想到的原因"] },
  ].concat(roles.map(function (role) {
    return {
      height: 1500,
      cells: [
        { text: role[0], bold: true, align: AlignmentType.CENTER },
        role[1],
        role[2],
        room(" "),
      ],
    };
  }))));
  return studentDoc(children);
}

function claimDoc() {
  const children = [];
  children.push.apply(children, sheetHead("我的主張"));
  children.push(dataTable([2400, CW - 2400], [
    { cells: [{ text: "問題", bold: true, fill: "E7EFE8", align: AlignmentType.CENTER }, "這些動物為什麼瀕臨絕種？"] },
    { cells: [{ text: "何時寫", bold: true, fill: "E7EFE8", align: AlignmentType.CENTER }, "打開閱讀之前。寫好後，這張紙留到第四課。"] },
    { cells: [{ text: "句式", bold: true, fill: "E7EFE8", align: AlignmentType.CENTER }, "我主張（動物）瀕臨絕種，是因為（原因）。"] },
  ]));
  children.push(p(" ", 16, { before: 60, after: 20 }));
  children.push(dataTable([1800, 4333, 4333], [
    { header: true, cells: ["", "動物", "是因為"] },
    { height: 1800, cells: [{ text: "主張一", bold: true, align: AlignmentType.CENTER }, room(" "), room(" ")] },
    { height: 1800, cells: [{ text: "主張二", bold: true, align: AlignmentType.CENTER }, room(" "), room(" ")] },
  ]));
  children.push(p(" ", 16, { before: 60, after: 20 }));
  children.push(dataTable([3200, CW - 3200], [
    { header: true, cells: ["聽老師說", "怎樣填"] },
    { cells: ["要寫兩個原因", "主張一和主張二都寫。兩個原因要不同。"] },
    { cells: ["只寫一個原因", "只寫主張一。主張二留空。"] },
  ]));
  children.push(p(" ", 16, { before: 60, after: 20 }));
  children.push(dataTable([7466, 1500, 1500], [
    { header: true, cells: ["我檢查", "是", "還未是"] },
    { height: 700, cells: ["主張回答了「為什麼瀕臨絕種」", { text: "□", align: AlignmentType.CENTER }, { text: "□", align: AlignmentType.CENTER }] },
    { height: 700, cells: ["若寫了主張二，它是另一個原因", { text: "□", align: AlignmentType.CENTER }, { text: "□", align: AlignmentType.CENTER }] },
  ]));
  return studentDoc(children);
}

function kwlDoc() {
  const children = [];
  children.push.apply(children, sheetHead("科學閱讀"));
  const kwCols = [3200, 2200, CW - 5400];
  children.push(dataTable(kwCols, [
    { header: true, cells: ["欄", "何時填", "寫什麼"] },
    { cells: ["K　我知道什麼", "打開課文之前", "遊戲、影片或角色扮演裏已經知道的。不肯定的也先寫下。"] },
    { cells: ["W　我想知道什麼", "打開課文之前", "寫成問題。例如：海龜為什麼變少？"] },
    { cells: ["L　我讀到什麼", "讀完之後", "只寫課文裏的句子，並寫上動物名稱。"] },
  ]));
  children.push(p(" ", 16, { before: 40, after: 20 }));
  children.push(dataTable([5233, 5233], [
    { header: true, cells: ["分組閱讀", "老師帶讀"] },
    {
      cells: [
        "聽老師分配。一組只讀一種動物，該動物的全部版本都要讀。可以有多於一組讀同一種。",
        "三種動物都讀。每種先看特徵，再看面臨的威脅。",
      ],
    },
  ]));
  children.push(p(" ", 16, { before: 80, after: 20 }));
  const cols = [3488, 3489, 3489];
  children.push(dataTable(cols, [
    { header: true, cells: ["K　我知道什麼", "W　我想知道什麼", "L　我讀到什麼"] },
    {
      cells: [
        { text: "讀文之前", fill: "F4F7F4", align: AlignmentType.CENTER, size: 28 },
        { text: "寫成問題", fill: "F4F7F4", align: AlignmentType.CENTER, size: 28 },
        { text: "讀完才寫", fill: "F4F7F4", align: AlignmentType.CENTER, size: 28 },
      ],
    },
    { height: 1600, cells: [room(" "), room(" "), room(" ")] },
    { height: 1600, cells: [room(" "), room(" "), room(" ")] },
    { height: 1600, cells: [room(" "), room(" "), room(" ")] },
  ]));

  children.push(p("讀的時候，畫出人類令動物變少的句子。然後填這張表。", 32, { breakBefore: true, before: 0, after: 60 }));
  children.push(dataTable([1400, 5466, 3600], [
    { header: true, cells: ["動物", "抄一句：人類令牠變少", "和我的主張"] },
    { height: 1600, cells: ["海龜", room(" "), { text: "有關／無關", align: AlignmentType.CENTER }] },
    { height: 1600, cells: ["老虎", room(" "), { text: "有關／無關", align: AlignmentType.CENTER }] },
    { height: 1600, cells: ["大象", room(" "), { text: "有關／無關", align: AlignmentType.CENTER }] },
  ]));
  children.push(p("分組閱讀的，先填自己那一種，另外兩種聽匯報後各抄一句。老師帶讀的，三種都填。", 32, { before: 80, after: 40 }));

  READINGS.forEach(function (reading) {
    children.push(p(reading.animal, 36, {
      breakBefore: true, before: 0, after: 60,
      align: AlignmentType.CENTER, run: { bold: true },
    }));
    reading.versions.forEach(function (version) {
      if (version.note) children.push(p(version.note, 24, { before: 0, after: 40 }));
      const paras = version.paras.map(function (para, index) {
        return p(para, 32, { before: index === 0 ? 40 : 80, after: 40, line: 400 });
      });
      children.push(dataTable([2400, CW - 2400], [{
        cantSplit: false,
        cells: [
          { text: version.title, bold: true, fill: "E7EFE8", align: AlignmentType.CENTER, valign: VerticalAlign.TOP, size: 28 },
          { text: paras, valign: VerticalAlign.TOP },
        ],
      }]));
      children.push(p(" ", 12, { before: 20, after: 20 }));
    });
  });
  return studentDoc(children);
}

function cerBlock(label, note) {
  return dataTable([2200, CW - 2200], [
    { header: true, cells: [label, note] },
    { height: 800, cells: [{ text: "判斷", bold: true }, "文本支持　　／　　文本不支持"] },
    { height: 1600, cells: [{ text: "證據", bold: true }, room("從課文抄一句，並寫上動物名稱。")] },
    {
      height: 1800,
      cells: [
        { text: "解釋", bold: true },
        lines([
          "這句證據_____我的主張，因為_____。",
          "若圈了「文本不支持」，寫出課文實際的原因。",
        ], 32),
      ],
    },
  ]);
}

function writeDoc() {
  const children = [];
  children.push.apply(children, sheetHead("主張、證據和解釋"));
  children.push(dataTable([2400, CW - 2400], [
    { cells: [{ text: "何時用", bold: true, fill: "E7EFE8", align: AlignmentType.CENTER }, "第四課。先把主張紙上的原句抄下來，再用課文判斷。"] },
  ]));
  children.push(p("一、我認識的瀕危物種", 32, { before: 160, after: 40, run: { bold: true } }));
  children.push(p("每種寫一項特徵，以及人類令牠變少的一個原因。", 32, { before: 0, after: 40 }));
  children.push(dataTable([1600, 4433, 4433], [
    { header: true, cells: ["動物", "一項特徵", "人類令牠變少的一個原因"] },
    { height: 1400, cells: ["海龜", room(" "), room(" ")] },
    { height: 1400, cells: ["老虎", room(" "), room(" ")] },
    { height: 1400, cells: ["大象", room(" "), room(" ")] },
  ]));
  children.push(p("分組閱讀的，把自己讀的那種寫清楚，另外兩種寫匯報聽到的。老師帶讀的，三種都寫。", 32, { before: 60, after: 40 }));

  children.push(p("二、抄回我的主張", 32, { before: 160, after: 40, run: { bold: true }, keepNext: true }));
  children.push(p("照抄主張紙，不要改成課文的句子。", 32, { before: 0, after: 40, keepNext: true }));
  children.push(dataTable([1800, 4333, 4333], [
    { header: true, cells: ["", "動物", "是因為"] },
    { height: 1200, cells: [{ text: "主張一", bold: true, align: AlignmentType.CENTER }, room(" "), room(" ")] },
    { height: 1200, cells: [{ text: "主張二", bold: true, align: AlignmentType.CENTER }, room("只寫一個主張的，這裏留空。"), room(" ")] },
  ]));

  children.push(p("三、用課文判斷", 32, { before: 200, after: 40, run: { bold: true }, keepNext: true }));
  children.push(p("只寫一個主張的，只做主張一。", 32, { before: 0, after: 60, keepNext: true }));
  children.push(cerBlock("主張一", "用課文判斷"));
  children.push(p(" ", 16, { before: 80, after: 20 }));
  children.push(cerBlock("主張二", "只寫一個主張的，整格留空"));

  children.push(p("四、我檢查", 32, { before: 200, after: 40, run: { bold: true }, keepNext: true }));
  children.push(dataTable([8466, 2000], [
    { header: true, cells: ["我檢查", "做到了"] },
    { height: 700, cells: ["主張回答了「為什麼瀕臨絕種」", { text: "□", align: AlignmentType.CENTER }] },
    { height: 700, cells: ["證據來自課文，並寫了動物名稱", { text: "□", align: AlignmentType.CENTER }] },
    { height: 700, cells: ["解釋說明這句證據怎樣和主張有關", { text: "□", align: AlignmentType.CENTER }] },
  ]));
  return studentDoc(children);
}

async function save(doc, name) {
  const buf = await Packer.toBuffer(doc);
  const dest = path.join(OUT, name);
  fs.writeFileSync(dest, buf);
  console.log(name, buf.length);
}

async function main() {
  await save(teacherDoc(), "6LD_教師課節流程.docx");
  await save(doDoc(), "6LD_動手記錄.docx");
  await save(claimDoc(), "6LD_主張紙.docx");
  await save(kwlDoc(), "6LD_KWL科學閱讀.docx");
  await save(writeDoc(), "6LD_寫作工作紙.docx");
  const oldNotes = path.join(OUT, "6LD_影片筆記.docx");
  if (fs.existsSync(oldNotes)) fs.unlinkSync(oldNotes);
}

main().catch(function (err) {
  console.error(err);
  process.exit(1);
});
