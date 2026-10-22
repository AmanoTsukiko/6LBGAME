const pptxgen = require("C:/Users/chleung/tmp_6lb2/ppt/node_modules/pptxgenjs");
const path = require("path");

const FONT = "Microsoft JhengHei UI";
const W = 13.333;
const H = 7.5;
const C = {
  cream: "F7F1E5",
  ink: "1A1F1C",
  muted: "5E675F",
  white: "FFFDF8",
  line: "E4D9C4",
  green: "2F5D3F",
  talk: "8C3A4B",
  read: "234E70",
  write: "6B4F2A",
  sand: "EFE6D4",
};

const PHASE = {
  "Do-it": C.green,
  "Talk-it": C.talk,
  "Read-it": C.read,
  "Write-it": C.write,
};

function shadow() {
  return { type: "outer", color: "000000", blur: 10, offset: 2, angle: 135, opacity: 0.08 };
}

function addChrome(slide, pres, label, page, total, phase) {
  slide.background = { color: C.cream };
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.16, h: H,
    fill: { color: PHASE[phase] || C.green }, line: { color: PHASE[phase] || C.green },
  });
  slide.addText(label, {
    x: 0.48, y: 0.22, w: 9.2, h: 0.32,
    fontFace: FONT, fontSize: 13, color: C.muted, margin: 0,
  });
  slide.addText(String(page).padStart(2, "0") + "  /  " + String(total).padStart(2, "0"), {
    x: W - 2.2, y: 0.22, w: 1.5, h: 0.32,
    fontFace: FONT, fontSize: 13, color: C.muted, align: "right", margin: 0,
  });
}

function addTitle(slide, text) {
  slide.addText(text, {
    x: 0.48, y: 0.64, w: 12.3, h: 0.52,
    fontFace: FONT, fontSize: 28, bold: true, color: C.ink, margin: 0,
  });
}

function card(slide, pres, x, y, w, h, title, lines, color) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.white },
    line: { color: C.line },
    shadow: shadow(),
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.1, h,
    fill: { color }, line: { color },
  });
  slide.addText(title, {
    x: x + 0.28, y: y + 0.14, w: w - 0.42, h: 0.36,
    fontFace: FONT, fontSize: 16, bold: true, color, margin: 0,
  });
  slide.addText(lines.map(function (line, i) {
    return { text: line, options: { bullet: false, breakLine: i < lines.length - 1 } };
  }), {
    x: x + 0.28, y: y + 0.54, w: w - 0.46, h: h - 0.68,
    fontFace: FONT, fontSize: 15, color: C.ink, margin: 0, valign: "top",
  });
}

function footer(slide, text) {
  slide.addText(text, {
    x: 0.48, y: H - 0.38, w: 12.3, h: 0.24,
    fontFace: FONT, fontSize: 12, color: C.muted, margin: 0,
  });
}

function flowBar(slide, pres, current) {
  const items = ["Do-it", "Talk-it", "Read-it", "Write-it"];
  const labels = ["Do-it  做", "Talk-it  說", "Read-it  讀", "Write-it  寫"];
  const gap = 0.16;
  const boxW = (12.3 - gap * 3) / 4;
  items.forEach(function (item, i) {
    const x = 0.48 + i * (boxW + gap);
    const on = item === current;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.28, w: boxW, h: 0.48,
      fill: { color: on ? PHASE[item] : C.sand },
      line: { color: on ? PHASE[item] : C.sand },
      rectRadius: 0.08,
    });
    slide.addText(labels[i], {
      x, y: 1.28, w: boxW, h: 0.48,
      fontFace: FONT, fontSize: 14, bold: true,
      color: on ? "FFFFFF" : C.ink, align: "center", valign: "middle", margin: 0,
    });
  });
}

const VIDEOS = [
  ["海龜", "Sea Turtles 101 | National Geographic", "https://www.youtube.com/watch?v=5Rmv3nliwCs", "全片 3 分 45 秒。邊看邊記。"],
  ["老虎", "What's Driving Tigers Toward Extinction? | National Geographic", "https://www.youtube.com/watch?v=qhrT0v7QSvc", "播 0:00–2:40。捕獵和虎製品。"],
  ["大象", "Why Elephants May Go Extinct in Your Lifetime | National Geographic", "https://www.youtube.com/watch?v=zJbFthhZRHY", "全片 1 分 52 秒。象牙和捕獵。"],
];

function build(kind) {
  const strong = kind === "strong";
  const pres = new pptxgen();
  pres.defineLayout({ name: "WIDE", width: W, height: H });
  pres.layout = "WIDE";
  pres.title = strong ? "6LD 瀕危物種 能力較高班" : "6LD 瀕危物種 能力稍遜班";
  pres.author = "佛教黃焯菴小學 科學科";
  pres.subject = "6LD1 認識一些瀕危物種　6LD2 解釋瀕危物種瀕臨絕種的原因";
  const klass = strong ? "能力較高班" : "能力稍遜班";
  const file = strong ? "6LD_能力較高班_課件.pptx" : "6LD_能力稍遜班_課件.pptx";
  const foot = "6LD 瀕危物種  ·  " + klass;

  const slides = [];

  function add(phase, title, draw, notes) {
    slides.push({ phase, title, draw, notes });
  }

  add(null, "封面", function (slide) {
    slide.background = { color: C.green };
    slide.addText("佛教黃焯菴小學　科學科　六年級", {
      x: 0.7, y: 1.15, w: 11.5, h: 0.4,
      fontFace: FONT, fontSize: 16, color: "D7E6DC", margin: 0,
    });
    slide.addText("瀕危物種", {
      x: 0.7, y: 1.7, w: 11.5, h: 0.9,
      fontFace: FONT, fontSize: 48, bold: true, color: "FFFFFF", margin: 0,
    });
    slide.addText(klass, {
      x: 0.7, y: 2.65, w: 11.5, h: 0.45,
      fontFace: FONT, fontSize: 24, color: "F3E6C8", margin: 0,
    });
    slide.addText("6LD1 認識一些瀕危物種", {
      x: 0.7, y: 3.5, w: 11, h: 0.38,
      fontFace: FONT, fontSize: 20, color: "FFFFFF", margin: 0,
    });
    slide.addText("6LD2 解釋瀕危物種瀕臨絕種的原因", {
      x: 0.7, y: 3.95, w: 11, h: 0.38,
      fontFace: FONT, fontSize: 20, color: "FFFFFF", margin: 0,
    });
    slide.addText("順序：Do-it → Talk-it → Read-it → Write-it", {
      x: 0.7, y: 4.7, w: 11, h: 0.4,
      fontFace: FONT, fontSize: 18, color: "F3E6C8", margin: 0,
    });
    slide.addText("Seeds of Science/Roots of Reading　·　KWL（Ogle, 1986）", {
      x: 0.7, y: 6.5, w: 11.5, h: 0.32,
      fontFace: FONT, fontSize: 14, color: "D7E6DC", margin: 0,
    });
  }, "開場告訴學生：今天先做、再說、然後讀、最後寫。主張在打開課文之前寫好。");

  add("Do-it", "跟著這四步，不要調換", function (slide, pres, page, total) {
    addChrome(slide, pres, klass + "  ·  課節地圖", page, total, "Do-it");
    addTitle(slide, "四步的順序");
    const steps = [
      ["Do-it", "第一課　先遇見數目下降", strong ? "玩捕獵遊戲，再邊看影片邊做筆記。" : "全班一起玩捕獵遊戲，再角色扮演一種瀕危物種。"],
      ["Talk-it", strong ? "第二課　先寫至少兩個主張" : "第二課　先寫一個主張", strong ? "討論可能的原因。兩個主張要是兩個原因。" : "用一句話說出動物為什麼瀕臨絕種。"],
      ["Read-it", "第三課　用 KWL 讀原文", strong ? "每一組只讀一種動物。多於一組可以讀同一種。" : "老師帶領，三種動物都讀。"],
      ["Write-it", "第四課　用文本檢驗主張", "寫認識的動物，並寫主張有沒有得到閱讀證據支持。"],
    ];
    steps.forEach(function (step, i) {
      const y = 1.4 + i * 1.3;
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.48, y, w: 12.3, h: 1.15,
        fill: { color: C.white }, line: { color: C.line }, rectRadius: 0.08,
      });
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.68, y: y + 0.28, w: 2.3, h: 0.58,
        fill: { color: PHASE[step[0]] }, line: { color: PHASE[step[0]] }, rectRadius: 0.08,
      });
      slide.addText(step[0], {
        x: 0.68, y: y + 0.28, w: 2.3, h: 0.58,
        fontFace: FONT, fontSize: 16, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0,
      });
      slide.addText(step[1], {
        x: 3.2, y: y + 0.14, w: 9.2, h: 0.4,
        fontFace: FONT, fontSize: 18, bold: true, color: C.ink, margin: 0,
      });
      slide.addText(step[2], {
        x: 3.2, y: y + 0.56, w: 9.2, h: 0.42,
        fontFace: FONT, fontSize: 15, color: C.muted, margin: 0,
      });
    });
  }, "四課，每課一步。第二課的主張紙留到第四課。若只有三課，第四課接在第三課讀完之後。");

  add("Do-it", "學習目標", function (slide, pres, page, total) {
    addChrome(slide, pres, klass, page, total, "Do-it");
    addTitle(slide, "這一課要學懂的");
    card(slide, pres, 0.48, 1.5, 6.0, 3.55, "6LD1", [
      "認識一些瀕危物種。",
      "",
      "海龜、老虎、大象：",
      "牠是什麼樣的動物，",
      "住在哪裏，靠什麼過活。",
    ], C.green);
    card(slide, pres, 6.85, 1.5, 6.0, 3.55, "6LD2", [
      "解釋瀕危物種瀕臨絕種的原因。",
      "",
      "原因要指到人做了什麼：",
      "捕獵、製品、棲息地。",
      "再用文本核對自己的主張。",
    ], C.read);
  }, "目標留在黑板一角，寫作時學生還看得見。");

  add("Do-it", "開課前", function (slide, pres, page, total) {
    addChrome(slide, pres, klass + "  ·  課前", page, total, "Do-it");
    addTitle(slide, "進教室前先放好");
    card(slide, pres, 0.48, 1.5, 6.0, 4.15, "老師做", [
      "第一課發動手記錄。",
      "第二課才發主張紙。",
      "第三課、主張寫好後",
      "才發 KWL。原文在後面。",
      "第四課才發寫作工作紙。",
      strong ? "分組留到第三課：" : "不按動物分組。",
      strong ? "一組一種，可多組同一種。" : "三種都由老師帶讀。",
    ], C.green);
    card(slide, pres, 6.85, 1.5, 6.0, 4.15, "學生這時還不用做", [
      "桌上還沒有課文。",
      "第一課填動手記錄。",
      "第二課在主張紙寫主張。",
      "主張紙留到第四課抄回。",
    ], C.write);
  }, "KWL 科學閱讀的後面就是三篇原文。第三課、主張寫好之後才發。主張寫在主張紙。寫作工作紙留到第四課，再把主張原句抄上去。若沒有印主張紙，可先寫便利貼，仍然不要發 KWL。");

  add("Do-it", "Do-it　捕獵遊戲", function (slide, pres, page, total) {
    addChrome(slide, pres, klass + "  ·  Do-it", page, total, "Do-it");
    addTitle(slide, "模擬：你是獵人，射中才少一隻");
    flowBar(slide, pres, "Do-it");
    card(slide, pres, 0.48, 1.95, 6.0, 4.15, "老師做", [
      "用瀏覽器打開",
      "hunting-game.html。",
      "讓全班看見開始的隻數：",
      "海龜 8、老虎 6、大象 6。",
      "學生是獵人，對準動物按一下。",
      "射中才少一隻。沒有射中，數目不變。",
      "每射中一次，停下，齊數還剩多少。",
      "問：數目怎樣變？先聽完再繼續。",
    ], C.green);
    card(slide, pres, 6.85, 1.95, 6.0, 4.15, "學生做", [
      "你是獵人。對準動物按一下。",
      "射中才少一隻。",
      "沒有射中，數目不變。",
      "在動手記錄寫下",
      "射中後還剩多少，和沒有射中時。",
      "重新開始會回到原數。",
      "再打開檔案也是一樣。",
    ], C.green);
  }, "可以說：我們先當獵人，看發生什麼事。射中才少一隻，沒有射中數目不變。每射中一次，數一數還剩多少。原因等一下再說。");

  if (strong) {
    add("Do-it", "Do-it　邊看影片邊做筆記", function (slide, pres, page, total) {
      addChrome(slide, pres, klass + "  ·  Do-it", page, total, "Do-it");
      addTitle(slide, "三條影片，看完一條先寫");
      VIDEOS.forEach(function (video, i) {
        const y = 1.4 + i * 1.7;
        slide.addShape(pres.shapes.RECTANGLE, {
          x: 0.48, y, w: 12.3, h: 1.55,
          fill: { color: C.white }, line: { color: C.line }, shadow: shadow(),
        });
        slide.addText(video[0], {
          x: 0.7, y: y + 0.16, w: 1.5, h: 0.4,
          fontFace: FONT, fontSize: 18, bold: true, color: C.green, margin: 0,
        });
        slide.addText(video[1], {
          x: 2.3, y: y + 0.14, w: 10.1, h: 0.38,
          fontFace: FONT, fontSize: 16, bold: true, color: "234E70", margin: 0,
          hyperlink: { url: video[2] },
        });
        slide.addText(video[2], {
          x: 2.3, y: y + 0.54, w: 10.1, h: 0.32,
          fontFace: FONT, fontSize: 13, color: C.muted, margin: 0,
          hyperlink: { url: video[2] },
        });
        slide.addText(video[3] + "  老師用廣東話說明，學生寫在動手記錄。", {
          x: 2.3, y: y + 0.96, w: 10.1, h: 0.36,
          fontFace: FONT, fontSize: 15, color: C.ink, margin: 0,
        });
      });
    }, "海龜若聽不清威脅，改播 Protecting marine turtles in Coastal Kenya | WWF UK  https://www.youtube.com/watch?v=_xv9p66beA0  漁網和產卵沙灘。老虎片沒有 When the buying stops, the killing can too. 不要為這一句另找片。");

    add("Do-it", "動手記錄記三樣", function (slide, pres, page, total) {
      addChrome(slide, pres, klass + "  ·  Do-it", page, total, "Do-it");
      addTitle(slide, "暫停的時候才寫");
      card(slide, pres, 0.48, 1.5, 4.0, 3.15, "1", ["這是什麼動物。", "我看見或聽見", "的一件事。"], C.green);
      card(slide, pres, 4.68, 1.5, 4.0, 3.15, "2", ["人類做了什麼，", "令這種動物變少。"], C.talk);
      card(slide, pres, 8.88, 1.5, 3.95, 3.15, "3", ["我還想知道", "的一個問題。", "留到 KWL 的 W。"], C.read);
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: 0.48, y: 4.9, w: 12.35, h: 0.85,
        fill: { color: C.sand }, line: { color: C.sand }, rectRadius: 0.08,
      });
      slide.addText("每條片約停 1 分鐘。寫在動手記錄的影片表。看完一條先寫，再播下一條。", {
        x: 0.72, y: 4.9, w: 11.9, h: 0.85,
        fontFace: FONT, fontSize: 16, color: C.ink, valign: "middle", margin: 0,
      });
    }, "每條片約停 1 分鐘。筆記不用完整句子。");
  } else {
    add("Do-it", "Do-it　角色扮演", function (slide, pres, page, total) {
      addChrome(slide, pres, klass + "  ·  Do-it", page, total, "Do-it");
      addTitle(slide, "每人演一種瀕危物種");
      const roles = [
        ["海龜", "雙手作划水", "我回到沙灘產卵。魚網困住我。人們拿走我的蛋，又把沙灘建成旅遊區。"],
        ["老虎", "雙手作爪", "森林被砍掉，我沒有地方住。人們捕獵我，取去皮和骨頭。"],
        ["大象", "手臂作象鼻", "人們佔去森林。他們捕獵我，只為了象牙。"],
      ];
      roles.forEach(function (role, i) {
        const x = 0.48 + i * 4.2;
        slide.addShape(pres.shapes.RECTANGLE, {
          x, y: 1.5, w: 4.0, h: 3.45,
          fill: { color: C.white }, line: { color: C.line }, shadow: shadow(),
        });
        slide.addShape(pres.shapes.RECTANGLE, {
          x, y: 1.5, w: 4.0, h: 0.7,
          fill: { color: C.green }, line: { color: C.green },
        });
        slide.addText(role[0], {
          x, y: 1.5, w: 4.0, h: 0.7,
          fontFace: FONT, fontSize: 22, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0,
        });
        slide.addText(role[1], {
          x: x + 0.22, y: 2.36, w: 3.56, h: 0.36,
          fontFace: FONT, fontSize: 15, color: C.talk, margin: 0,
        });
        slide.addText(role[2], {
          x: x + 0.22, y: 2.82, w: 3.56, h: 1.9,
          fontFace: FONT, fontSize: 16, color: C.ink, margin: 0,
        });
      });
    }, "三種都要有人演。學生跟著投影說。台詞方便開口，稍後仍以原文核對。請兩三位出來表演，其餘跟著說。");
  }

  add("Talk-it", strong ? "Talk-it　至少兩個主張" : "Talk-it　只寫一個主張", function (slide, pres, page, total) {
    addChrome(slide, pres, klass + "  ·  Talk-it", page, total, "Talk-it");
    addTitle(slide, strong ? "討論原因，寫下多於一個主張" : "全班討論後，每人一個主張");
    flowBar(slide, pres, "Talk-it");
    card(slide, pres, 0.48, 1.95, 6.0, 4.15, "老師做", strong ? [
      "問：動物瀕臨絕種，",
      "可能有什麼原因？",
      "先讓學生說，再寫。",
      "發主張紙。不要發 KWL。",
      "每人至少兩個主張，",
      "兩個是不同的原因。",
      "主張紙留到第四課。",
      "不要先讀課文。",
    ] : [
      "問：這些動物為什麼",
      "會瀕臨絕種？",
      "聽兩三個人說。",
      "發主張紙。每人只寫一個。",
      "主張二留空。",
      "KWL 和寫作工作紙未發。",
      "不要先讀課文。",
    ], C.talk);
    card(slide, pres, 6.85, 1.95, 6.0, 4.15, "學生做", strong ? [
      "和旁邊的人說出原因。",
      "在主張紙寫主張一",
      "和主張二。",
      "",
      "句式：",
      "我主張（動物）瀕臨絕種，",
      "是因為（原因）。",
    ] : [
      "只在主張紙寫主張一。",
      "主張二留空。",
      "",
      "句式：",
      "我主張這些動物瀕臨絕種，",
      "是因為______。",
    ], C.talk);
  }, strong
    ? "主張紙收好，第四課發還。若兩個主張其實是同一原因，請學生改第二個。例如捕獵和棲息地才算兩個。KWL 和寫作工作紙這時都不發。"
    : "主張紙收好，第四課發還。空著的，用角色扮演裏的一句幫他起頭，仍由他自己寫完。KWL 和寫作工作紙這時都不發。");

  add("Read-it", "Read-it　KWL", function (slide, pres, page, total) {
    addChrome(slide, pres, klass + "  ·  Read-it", page, total, "Read-it");
    addTitle(slide, "打開文本之前，先填 K 和 W");
    flowBar(slide, pres, "Read-it");
    const cols = [
      ["K  我知道什麼", "Know", "遊戲、影片或角色扮演裏已經知道的，都寫上。不肯定的也先留著。"],
      ["W  我想知道什麼", "Want to know", "寫成問題。例如：文本有沒有寫到我的主張？還有沒有別的原因？"],
      ["L  我讀到什麼", "Learned", "讀完才填。只寫文本裏的句子，並寫上動物名稱。"],
    ];
    cols.forEach(function (col, i) {
      const x = 0.48 + i * 4.2;
      slide.addShape(pres.shapes.RECTANGLE, {
        x, y: 1.95, w: 4.0, h: 2.85,
        fill: { color: C.white }, line: { color: C.line }, shadow: shadow(),
      });
      slide.addText(col[0], {
        x: x + 0.22, y: 2.1, w: 3.56, h: 0.42,
        fontFace: FONT, fontSize: 18, bold: true, color: C.read, margin: 0,
      });
      slide.addText(col[1], {
        x: x + 0.22, y: 2.52, w: 3.56, h: 0.28,
        fontFace: FONT, fontSize: 14, color: C.muted, margin: 0,
      });
      slide.addText(col[2], {
        x: x + 0.22, y: 2.88, w: 3.56, h: 1.7,
        fontFace: FONT, fontSize: 16, color: C.ink, margin: 0, valign: "top",
      });
    });
  }, "KWL（Ogle, 1986：Know, Want to know, Learned）。W 不要接受「老虎」這種課題詞，要改成問題。");

  add("Read-it", strong ? "每一組只讀一種動物" : "老師帶讀三種動物", function (slide, pres, page, total) {
    addChrome(slide, pres, klass + "  ·  Read-it", page, total, "Read-it");
    addTitle(slide, strong ? "分組閱讀" : "全班跟著老師讀");
    card(slide, pres, 0.48, 1.5, 6.0, 4.35, "老師做", strong ? [
      "主張已寫在主張紙。",
      "現在才派 KWL 科學閱讀。",
      "每一組只讀一種動物。",
      "海龜、老虎或大象，",
      "該動物的全部版本都要讀。",
      "可以有多於一組",
      "讀同一種動物。",
      "盡量三種都有組別，",
      "匯報時全班才聽得全。",
    ] : [
      "主張已寫在主張紙。",
      "現在才派 KWL 科學閱讀。",
      "不按動物分組。",
      "三種動物都帶領閱讀。",
      "每種先讀一兩句特徵，",
      "再讀「面臨的威脅」。",
      "讀到原因句，讓學生畫線。",
      "每讀完一種，在 L 寫",
      "一個原因。",
    ], C.read);
    card(slide, pres, 6.85, 1.5, 6.0, 4.35, "學生做", strong ? [
      "畫出人類令動物變少",
      "的句子。",
      "L 只寫自己那篇。",
      "匯報一句原因，",
      "並說出它在文本的哪裏。",
      "",
      "海龜：漁網、肉、蛋、沙灘、殼。",
      "老虎：森林、皮、骨、偷獵。",
      "大象：棲息地、象牙。",
    ] : [
      "手指跟著老師的句子。",
      "畫線：",
      "海龜的漁網、蛋、沙灘、殼；",
      "老虎的森林、皮、骨、偷獵；",
      "大象的棲息地和象牙。",
      "L 欄三種各寫一句。",
    ], C.read);
  }, strong
    ? "匯報只收一個人為原因加文本出處。老虎、大象提醒還有棲息地，海龜文本沒有寫森林消失。"
    : "帶讀句子見教師課節流程。跳過老虎英文那一題，告訴學生那句不在今天的影片裏。");

  add("Write-it", "Write-it　用證據檢驗主張", function (slide, pres, page, total) {
    addChrome(slide, pres, klass + "  ·  Write-it", page, total, "Write-it");
    addTitle(slide, "回到讀文之前的主張");
    flowBar(slide, pres, "Write-it");
    card(slide, pres, 0.48, 1.95, 6.0, 4.15, "老師做", [
      "現在才發寫作工作紙。",
      "把主張紙上的原句",
      "抄到主張表。",
      strong ? "自己讀的那種寫詳細，" : "三種各寫一項特徵",
      strong ? "另兩種寫匯報聽到的一句。" : "和一個原因。",
      strong ? "主張一和主張二都做" : "只做主張一。",
      strong ? "判斷、證據、解釋。" : "主張二整格留空。",
      "證據要寫動物名稱。",
    ], C.write);
    card(slide, pres, 6.85, 1.95, 6.0, 4.15, "學生做", [
      "主張保持原句，",
      "不要改成課文。",
      "圈：文本支持，",
      "或文本不支持。",
      "證據：抄課文一句。",
      "解釋：這句證據怎樣",
      "和主張有關。",
    ], C.write);
  }, "只寫「因為會絕種」不算完成。海龜文本不支持「只因為森林消失」。捕獵則三種文本都有證據。");

  add("Write-it", "四課怎樣分", function (slide, pres, page, total) {
    addChrome(slide, pres, klass + "  ·  時間", page, total, "Write-it");
    addTitle(slide, "一課做一步");
    const rows = strong
      ? [
        ["第一課", "Do-it", "捕獵遊戲，再看三條影片，填動手記錄。"],
        ["第二課", "Talk-it", "討論原因。主張紙寫至少兩個不同原因。"],
        ["第三課", "Read-it", "先填 K、W。分組讀一種動物，再填 L 和證據表。"],
        ["第四課", "Write-it", "抄回主張，用課文寫證據和解釋。"],
      ]
      : [
        ["第一課", "Do-it", "捕獵遊戲，再角色扮演，填動手記錄。"],
        ["第二課", "Talk-it", "討論原因。主張紙只寫一個主張。"],
        ["第三課", "Read-it", "先填 K、W。老師帶讀三種，再填 L 和證據表。"],
        ["第四課", "Write-it", "抄回主張一，用課文寫證據和解釋。"],
      ];
    rows.forEach(function (row, i) {
      const y = 1.42 + i * 1.15;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.48, y, w: 12.3, h: 1.02,
        fill: { color: C.white }, line: { color: C.line },
      });
      slide.addText(row[0], {
        x: 0.7, y, w: 2.0, h: 1.02,
        fontFace: FONT, fontSize: 16, bold: true, color: C.green, valign: "middle", margin: 0,
      });
      slide.addText(row[1], {
        x: 2.7, y, w: 2.0, h: 1.02,
        fontFace: FONT, fontSize: 16, bold: true, color: PHASE[row[1]], valign: "middle", margin: 0,
      });
      slide.addText(row[2], {
        x: 4.8, y, w: 7.7, h: 1.02,
        fontFace: FONT, fontSize: 16, color: C.ink, valign: "middle", margin: 0,
      });
    });
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.48, y: 6.22, w: 12.3, h: 0.62,
      fill: { color: C.sand }, line: { color: C.sand }, rectRadius: 0.08,
    });
    slide.addText("若只有三課，第四課接在第三課讀完之後。主張仍然在打開課文之前寫好。", {
      x: 0.7, y: 6.22, w: 11.9, h: 0.62,
      fontFace: FONT, fontSize: 15, color: C.ink, valign: "middle", margin: 0,
    });
  }, "若只有三課，讀完就接著寫，主張紙仍然先寫好。KWL 在主張之後才發。");

  add("Write-it", "原文照讀", function (slide, pres, page, total) {
    addChrome(slide, pres, klass + "  ·  守住原文", page, total, "Write-it");
    addTitle(slide, "數字和字句不要換成新的");
    card(slide, pres, 0.48, 1.5, 12.35, 4.15, "老師知道即可，不要在學生課文上改", [
      "簡介已改成短句。事實仍保留。",
      "面臨威脅的段落是原文。照讀。",
      "海龜文仍寫小海龜兩日後破殼而出。",
      "老虎用「本世紀初」、約十萬頭、餘下約五千頭。",
      "大象寫 1970 年至 1989 年由 150 萬頭減至不足 60 萬頭。",
      "這些數字不要換成較新的數目。",
      "英文句 When the buying stops, the killing can too.",
      "不在今天的影片裏。不用找這一句。",
    ], C.write);
  }, "這些數字可能過時。這一課的閱讀任務是用眼前的文本做證據，不是更新資料。");

  add("Write-it", "文獻", function (slide, pres, page, total) {
    addChrome(slide, pres, klass + "  ·  文獻", page, total, "Write-it");
    addTitle(slide, "這一課依據的方法");
    card(slide, pres, 0.48, 1.38, 12.35, 1.72, "Seeds of Science/Roots of Reading", [
      "Lawrence Hall of Science。Do-it、Talk-it、Read-it、Write-it。",
      "firsthand 是模擬、觀察和扮演。secondhand 是閱讀並判斷原文。",
      "Cervetti, Pearson, Barber, Hiebert, & Bravo（2007）。",
    ], C.green);
    card(slide, pres, 0.48, 3.42, 12.35, 1.4, "KWL（Ogle, 1986）", [
      "Know, Want to know, Learned。W 寫問題，L 只寫課文。",
      "Ogle, D. M. (1986). The Reading Teacher, 39(6), 564–570.",
    ], C.read);
    card(slide, pres, 0.48, 5.14, 12.35, 1.58, "主張、證據、解釋", [
      "Claim, Evidence, Reasoning。McNeill & Krajcik（2012）。",
      "主張在讀文前寫下。證據抄課文一句，並寫動物名稱。",
      "解釋說明這句證據為什麼支持或不支持主張。",
    ], C.write);
    footer(slide, "教師課節流程見同一資料夾的 6LD_教師課節流程.docx");
  }, "學生工作紙不抄這些文獻。同事若要查方法，看這一頁和教師課節流程。");

  const total = slides.length;
  slides.forEach(function (item, index) {
    const slide = pres.addSlide();
    item.draw(slide, pres, index + 1, total);
    if (item.notes) slide.addNotes(item.notes);
  });

  // Fix page numbers: chrome was called with 0,0. Re-loop is hard after draw.
  // Page numbers were placeholders. Set them by redrawing? The chrome calls used 0,0.
  // I'll pass index into draw via closure by rebuilding... already drawn.
  // Patch: include page in add() by drawing with known total since slides array is complete before addSlide.
  return pres.writeFile({ fileName: path.join(__dirname, file) }).then(function () {
    console.log(file, total);
  });
}

// The page numbers were written as 00 because draw runs before we know... 
// actually total is known before forEach, but addChrome inside draw used 0,0.
// Rebuild properly: change add() to not draw until the end — already the case,
// and draw closes over nothing for page. I need to pass index.
// Easiest fix: edit addChrome calls... they're inside draw with literal 0,0.
// I'll update the draw functions' addChrome to use arguments.
// Looking at the code, draw is function (slide, pres, page, total) and forEach calls item.draw(slide, pres, index+1, total)
// but the functions ignore the extra args and hardcode 0,0.
// I'll fix by a second pass in the file... let me change the forEach to stamp page numbers
// using a known pattern. Simpler to edit addChrome invocation.
// I'll post-process: since I haven't run yet, I'll fix the source so each draw accepts page, total.

build("strong").then(function () { return build("weak"); }).catch(function (err) {
  console.error(err);
  process.exit(1);
});
