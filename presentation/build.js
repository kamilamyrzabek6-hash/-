const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/af42125f-92fe-4641-a239-6abd1d9556c6_f757a656-ca4c-4235-ab13-8eb81086f4ce/pptx/scripts/apply_theme.js");

const OUT = "/home/user/-/presentation/Slater_Kuzyrettilik_ZhI_Myrzabek.pptx";

const THEME = {
  name: "Atom Teal",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2430", lt1: "FFFFFF", dk2: "0E3B47", lt2: "EAF3F5",
    accent1: "0F7C8C", accent2: "F2A541", accent3: "D1495B",
    accent4: "5B8E7D", accent5: "3D5A80", accent6: "8D99AE",
    hlink: "0F7C8C", folHlink: "3D5A80",
  },
};
const H = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Химиялық білім берудегі құзыреттілік тәсіл және жасанды интеллект";
pres.author = "Мырзабек Камила";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const W = 13.333;

// ---------- icons ----------
async function icon(Comp, color = "FFFFFF", size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- layouts ----------
const FOOT = "Мырзабек К. · Химиялық білім берудегі құзыреттілік тәсіл және ЖИ";
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 1.35, w: 8.6, h: 2.9, fontFace: THEME.headFontFace, fontSize: 32, align: "left", bold: true, color: C.background1, valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.7, y: 4.45, w: 8.6, h: 1.0, fontSize: 17, align: "left", color: C.background2, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.85, fontFace: THEME.headFontFace, fontSize: 32, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.6, y: 7.05, w: 9, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: C.accent6, align: "right" },
});

// ---------- helpers ----------
function card(s, x, y, w, h, fill = C.background2, name = "card") {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.12, objectName: name });
}
function circleIcon(s, data, x, y, d, fill) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, objectName: "icon-circle" });
  const p = d * 0.22;
  s.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: "icon" });
}
function txt(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, fontSize: 16, valign: "top" }, o)); }
function atomRings(s, cx, cy, r, color, transparency = 60) {
  // decorative electron shells motif
  [1, 0.72, 0.44].forEach((k, i) => {
    s.addShape(pres.shapes.OVAL, { x: cx - r * k, y: cy - r * k, w: 2 * r * k, h: 2 * r * k, fill: { type: "none" }, line: { color, width: 1.5, transparency }, objectName: "shell" });
  });
  s.addShape(pres.shapes.OVAL, { x: cx - r * 0.12, y: cy - r * 0.12, w: r * 0.24, h: r * 0.24, fill: { color: C.accent2 }, line: { color: C.accent2 }, objectName: "nucleus" });
  [[1, 30], [0.72, 200], [0.44, 120], [1, 250]].forEach(([k, deg]) => {
    const a = (deg * Math.PI) / 180, e = r * 0.06;
    s.addShape(pres.shapes.OVAL, { x: cx + r * k * Math.cos(a) - e, y: cy + r * k * Math.sin(a) - e, w: 2 * e, h: 2 * e, fill: { color: C.accent1 }, line: { color: C.background1, width: 1 }, objectName: "electron" });
  });
}
const chartBase = () => ({
  catAxisLabelColor: "1B2430", valAxisLabelColor: "1B2430", catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, valGridLine: { color: "D5E2E6", size: 0.75 }, catGridLine: { style: "none" },
  titleFontFace: "+mn-lt", titleFontSize: 14, titleColor: "0E3B47", showTitle: true,
  legendFontFace: "+mn-lt", legendFontSize: 12, legendColor: "1B2430",
  dataLabelFontFace: "+mn-lt", dataLabelFontSize: 10, dataLabelColor: "1B2430",
});

(async () => {
  const I = {};
  const list = { atom: fa.FaAtom, brain: fa.FaBrain, robot: fa.FaRobot, bulb: fa.FaLightbulb, flask: fa.FaFlask, chart: fa.FaChartLine,
    search: fa.FaSearch, scale: fa.FaBalanceScale, comments: fa.FaComments, grad: fa.FaUserGraduate, check: fa.FaCheckCircle,
    warn: fa.FaExclamationTriangle, question: fa.FaQuestion, puzzle: fa.FaPuzzlePiece, laptop: fa.FaLaptopCode, book: fa.FaBookOpen,
    sync: fa.FaSyncAlt, target: fa.FaBullseye, layers: fa.FaLayerGroup, flag: fa.FaFlagCheckered, eye: fa.FaEye, key: fa.FaKey, users: fa.FaUsers };
  for (const [k, v] of Object.entries(list)) I[k] = await icon(v);

  // ================= 1. TITLE =================
  pres.addSection({ title: "Кіріспе" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Кіріспе" });
  txt(s, "СЕМИНАР · ХИМИЯЛЫҚ БІЛІМ БЕРУДЕГІ ҚҰЗЫРЕТТІЛІК ТӘСІЛ ЖӘНЕ ЖАСАНДЫ ИНТЕЛЛЕКТ", { x: 0.7, y: 0.8, w: 9, h: 0.4, fontSize: 13, bold: true, color: C.accent2, charSpacing: 1 });
  s.addText("Атом құрылысы мен Периодтық заңды Слэйтер ережесі арқылы проблемалық оқытуда ЖИ құралдарын кіріктіру", { placeholder: "title" });
  s.addText("Диссертация тақырыбы: «Бейорганикалық химияны проблемалық оқытуда ЖИ құралдарын кіріктіру арқылы студенттердің сыни ойлауын дамыту»", { placeholder: "body" });
  atomRings(s, 11.1, 3.3, 1.75, "FFFFFF", 40);
  txt(s, [
    { text: "Орындаған: ", options: { bold: true } }, { text: "Мырзабек Камила, 7M01515 – Химия, 2-курс магистранты", options: { breakLine: true } },
    { text: "Ғылыми жетекшісі: ", options: { bold: true } }, { text: "т.ғ.к., профессор Тапалова Анипа Сейдалықызы", options: { breakLine: true } },
    { text: "Қорқыт Ата атындағы Қызылорда университеті · 2026", options: {} },
  ], { x: 0.7, y: 5.7, w: 9.5, h: 1.2, fontSize: 15, color: C.background1, paraSpaceAfter: 4 });
  s.addNotes("Сәлеметсіздер ме! Бүгінгі баяндамам диссертациямның негізгі идеясын «Химиялық білім берудегі құзыреттілік тәсіл және жасанды интеллект» пәнімен байланыстырады. Бейорганикалық химияның ішінен Атом құрылысы және Периодтық заң тақырыптарын таңдап, оларды Слэйтер ережесі арқылы проблемалық оқытуда ChatGPT және Claude құралдарын қолдану жолын көрсетемін. (≈1 мин)");

  // ================= 2. OUTLINE =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Баяндама жоспары", { placeholder: "title" });
  const outline = [
    ["01", "Өзектілігі және зерттеу аппараты", "Неге бұл мәселе маңызды, не зерттеледі"],
    ["02", "Теориялық негіздер", "Құзыреттілік тәсіл, сыни ойлау, проблемалық оқыту, ЖИ"],
    ["03", "Кіріктіру моделі", "Мақсат – мазмұн – әдіс – құрал – нәтиже"],
    ["04", "Слэйтер ережесі", "Z* есептеу, периодтылықты түсіндіру, модель шектеулері"],
    ["05", "Проблемалық жағдаяттар және ЖИ", "Fe, сілтілік металдар, ЖИ жауабын сыни бағалау"],
    ["06", "Сабақ үлгісі және бағалау", "50 минуттық сабақ, рубрика, күтілетін нәтижелер"],
  ];
  outline.forEach((o, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.15, y = 1.55 + row * 1.75;
    card(s, x, y, 5.95, 1.5, C.background2, "outline-card");
    txt(s, o[0], { x: x + 0.3, y: y + 0.3, w: 1.0, h: 0.9, fontSize: 36, bold: true, color: C.accent1, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, o[1], { x: x + 1.4, y: y + 0.28, w: 4.4, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    txt(s, o[2], { x: x + 1.4, y: y + 0.78, w: 4.4, h: 0.5, fontSize: 14, color: C.text1 });
  });
  s.addNotes("Баяндама алты бөлімнен тұрады. Алдымен өзектілікті және диссертацияның зерттеу аппаратын ұсынамын, содан кейін теориялық негіздерді қарастырамыз. Негізгі бөлім – Слэйтер ережесі арқылы құрылған проблемалық жағдаяттар және оларды ЖИ құралдарымен шешу. Соңында сабақ үлгісі мен бағалау рубрикасын көрсетемін. (≈1,5 мин)");

  // ================= 3. RELEVANCE =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Зерттеудің өзектілігі", { placeholder: "title" });
  const stats = [
    ["2024–2029", "ҚР Жасанды интеллектті дамыту тұжырымдамасы (ҚР Үкіметінің № 592 қаулысы) білім беруде ЖИ-ді кеңінен енгізуді көздейді", C.accent1],
    ["20% → 60%", "ЖИ пәндерін оқытатын ЖОО үлесін 2025 жылдан 2029 жылға дейін арттыру жоспарланған (Тұжырымдама, 2024)", C.accent5],
    ["Сенімді, бірақ қате", "ChatGPT химия есептеріне кейде қисынды көрінетін қате жауап береді, бұл жаңадан үйреніп жүрген студентті сендіре алады (Clark, 2023)", C.accent3],
  ];
  stats.forEach((st, i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.55, 3.9, 3.0, C.background2, "stat-card");
    txt(s, st[0], { x: x + 0.3, y: 1.8, w: 3.4, h: 0.9, fontSize: i === 2 ? 28 : 40, bold: true, color: st[2], fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, st[1], { x: x + 0.3, y: 2.85, w: 3.35, h: 1.6, fontSize: 14 });
  });
  circleIcon(s, I.question, 0.6, 5.0, 0.9, C.accent2);
  txt(s, [
    { text: "Қайшылық: ", options: { bold: true, color: C.text2 } },
    { text: "студенттер ЖИ құралдарын белсенді қолданады, бірақ олардың жауабын тексеру, дәлелдеу және модельдің шектеулерін бағалау дағдылары (сыни ойлау) жүйелі дамытылмайды. Ал «Атом құрылысы» тақырыбы көбіне ережелерді жаттаумен шектеледі." },
  ], { x: 1.75, y: 4.95, w: 10.9, h: 1.4, fontSize: 16, valign: "middle" });
  s.addNotes("Өзектілік үш дәлелге сүйенеді. Біріншіден, мемлекеттік саясат: 2024 жылы қабылданған Жасанды интеллектті дамытудың 2024–2029 жылдарға арналған тұжырымдамасы білім беруде ЖИ-ді енгізуді талап етеді. Екіншіден, ЖИ оқытатын ЖОО үлесін 20%-дан 60%-ға дейін арттыру жоспары бар. Үшіншіден, Clark (2023) зерттеуі ChatGPT-дің химия емтихан сұрақтарына қате, бірақ сенімді жауаптар беретінін көрсетті. Автордың өзі ЖИ жауаптарын студенттер сынап, жақсартатын тапсырмаларды ұсынады. Бұл менің зерттеуімнің негізгі идеясы. (≈3 мин)");

  // ================= 4. RESEARCH APPARATUS =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Зерттеу аппараты (ұсынылған нұсқа)", { placeholder: "title" });
  const app = [
    [I.target, "Мақсаты", "Бейорганикалық химияны проблемалық оқытуда ЖИ құралдарын кіріктіру арқылы студенттердің сыни ойлауын дамыту әдістемесін теориялық негіздеу және тәжірибе жүзінде тексеру"],
    [I.layers, "Нысаны", "Жоғары оқу орнында бейорганикалық химияны оқыту процесі"],
    [I.search, "Пәні", "«Атом құрылысы және Периодтық заң» тақырыбын Слэйтер ережесі негізінде, ChatGPT пен Claude көмегімен проблемалық оқыту әдістемесі"],
  ];
  app.forEach((a, i) => {
    const y = 1.45 + [0, 1.95, 3.3][i];
    circleIcon(s, a[0], 0.6, y + 0.1, 0.75, C.accent1);
    txt(s, a[1], { x: 1.55, y: y + 0.05, w: 4.6, h: 0.35, fontSize: 16, bold: true, color: C.text2 });
    txt(s, a[2], { x: 1.55, y: y + 0.42, w: 5.1, h: 1.4, fontSize: 14 });
  });
  card(s, 7.0, 1.5, 5.7, 5.25, C.background2, "tasks-card");
  txt(s, "Міндеттері", { x: 7.3, y: 1.7, w: 5.1, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
  txt(s, [
    { text: "Құзыреттілік тәсіл, сыни ойлау және проблемалық оқыту бойынша әдебиеттерге шолу жасау", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "ЖИ құралдарын проблемалық оқытуға кіріктіру моделін құру", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "Слэйтер ережесіне негізделген проблемалық жағдаяттар жүйесін әзірлеу", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "Сыни ойлауды бағалау рубрикасын құрастыру", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "Педагогикалық эксперимент жүргізіп, тиімділікті бағалау", options: { bullet: { type: "number" } } },
  ], { x: 7.3, y: 2.2, w: 5.15, h: 2.7, fontSize: 14, paraSpaceAfter: 6 });
  txt(s, [
    { text: "Болжам: ", options: { bold: true, color: C.accent1 } },
    { text: "егер ЖИ құралдары дайын жауап көзі ретінде емес, тексерілетін және сынға алынатын «серіктес» ретінде проблемалық жағдаяттарға кіріктірілсе, студенттердің сыни ойлау деңгейі артады." },
  ], { x: 7.3, y: 5.0, w: 5.15, h: 1.6, fontSize: 14 });
  s.addNotes("Зерттеу аппараты әлі ресми бекітілмегендіктен, мұнда ұсынылған нұсқасын көрсетіп отырмын. Пәні бүкіл бейорганикалық химия емес, нақты бір тақырып: Атом құрылысы және Периодтық заң, оны Слэйтер ережесімен байланыстырамыз. Болжамның негізгі идеясы: ЖИ дайын жауап беретін құрал емес, студент тексеретін және сынға алатын серіктес болуы керек. Қазір әдеби шолу кезеңіндемін, сондықтан 5-міндет алдағы жұмыс. (≈3 мин)");

  // ================= 5. COMPETENCY APPROACH =================
  pres.addSection({ title: "Теориялық негіздер" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Теориялық негіздер" });
  s.addText("Құзыреттілік тәсіл: білімнен әрекетке", { placeholder: "title" });
  // left: traditional, right: competency
  card(s, 0.6, 1.55, 5.4, 3.6, C.background2, "trad");
  txt(s, "Дәстүрлі (білімдік) тәсіл", { x: 0.9, y: 1.75, w: 4.8, h: 0.4, fontSize: 18, bold: true, color: C.accent6 });
  txt(s, [
    { text: "Нәтиже: «білім, білік, дағды»", options: { bullet: true, breakLine: true } },
    { text: "Слэйтер ережесін жаттап, формула бойынша есептеу", options: { bullet: true, breakLine: true } },
    { text: "Оқытушы – ақпарат көзі", options: { bullet: true, breakLine: true } },
    { text: "Бағалау: дұрыс жауап бар ма, жоқ па", options: { bullet: true } },
  ], { x: 0.9, y: 2.3, w: 4.9, h: 2.6, fontSize: 15, paraSpaceAfter: 8 });
  s.addShape(pres.shapes.RIGHT_ARROW, { x: 6.2, y: 2.95, w: 0.9, h: 0.8, fill: { color: C.accent2 }, line: { color: C.accent2 }, objectName: "arrow" });
  card(s, 7.3, 1.55, 5.4, 3.6, C.accent1, "comp");
  txt(s, "Құзыреттілік тәсіл", { x: 7.6, y: 1.75, w: 4.8, h: 0.4, fontSize: 18, bold: true, color: C.background1 });
  txt(s, [
    { text: "Нәтиже: білімді жаңа жағдаятта қолдану қабілеті", options: { bullet: true, breakLine: true } },
    { text: "Z* арқылы иондану энергиясы мен радиустың өзгерісін түсіндіру, болжау", options: { bullet: true, breakLine: true } },
    { text: "Оқытушы – ұйымдастырушы, тәлімгер", options: { bullet: true, breakLine: true } },
    { text: "Бағалау: дәлелдеу мен ойлау сапасы", options: { bullet: true } },
  ], { x: 7.6, y: 2.3, w: 4.9, h: 2.6, fontSize: 15, color: C.background1, paraSpaceAfter: 8 });
  // formula row
  const parts = [["Білім", C.accent5], ["Білік", C.accent5], ["Тәжірибе", C.accent5], ["Құндылық", C.accent5]];
  parts.forEach((p, i) => {
    const x = 0.6 + i * 2.35;
    card(s, x, 5.5, 2.0, 0.8, p[1], "comp-part");
    txt(s, p[0], { x, y: 5.5, w: 2.0, h: 0.8, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
    txt(s, i < 3 ? "+" : "=", { x: x + 2.0, y: 5.5, w: 0.35, h: 0.8, fontSize: 24, bold: true, color: C.text2, align: "center", valign: "middle" });
  });
  card(s, 10.0, 5.5, 2.7, 0.8, C.accent2, "comp-result");
  txt(s, "Құзыреттілік", { x: 10.0, y: 5.5, w: 2.7, h: 0.8, fontSize: 18, bold: true, color: C.text1, align: "center", valign: "middle" });
  txt(s, "Дереккөз: Хуторской А.В. (2003); Зимняя И.А. (2003); ҚР жоғары және жоғары оқу орнынан кейінгі білім берудің МЖМБС (2022)", { x: 0.6, y: 6.5, w: 12.1, h: 0.35, fontSize: 11, color: C.accent6 });
  s.addNotes("Құзыреттілік тәсілдің мәні: оқу нәтижесі тек білім мен білік емес, оны жаңа жағдаятта қолдана білу. А.В. Хуторской мен И.А. Зимняя «құзырет» пен «құзыреттілікті» ажыратады: құзырет – талап, нормативтік сипаттама, ал құзыреттілік – тұлғаның сол талапты орындау қабілеті. Біздің тақырыпқа қолдансақ: Слэйтер ережесін жаттау жеткіліксіз, студент Z* арқылы иондану энергиясы мен атом радиусының неге солай өзгеретінін түсіндіріп, болжай алуы керек. (≈3 мин)");

  // ================= 6. WHICH COMPETENCIES =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Теориялық негіздер" });
  s.addText("Қай құзыреттерді дамыту тиімді?", { placeholder: "title" });
  const rows = [
    [{ text: "Құзырет", options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 } } }, { text: "Тақырыптағы көрінісі", options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 } } }, { text: "Басымдық", options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 }, align: "center" } }],
    ["Пәндік (химиялық)", "Электрондық конфигурация, Z* есептеу, периодтық заңдылықтарды түсіндіру", { text: "Негізгі", options: { bold: true, color: H.accent1, align: "center" } }],
    ["Сыни ойлау", "Деректерді талдау, дәлелдеу, Слэйтер моделінің шектеулерін бағалау", { text: "Өзек", options: { bold: true, color: H.accent3, align: "center" } }],
    ["Цифрлық / ЖИ сауаттылығы", "Сұраныс (промпт) құрастыру, ЖИ жауабын тексеру, академиялық адалдық", { text: "Негізгі", options: { bold: true, color: H.accent1, align: "center" } }],
    ["Зерттеушілік", "Болжам ұсыну, есептеу арқылы тексеру, анықтамалық деректермен салыстыру", { text: "Қосымша", options: { color: H.accent5, align: "center" } }],
    ["Коммуникативтік", "Топта талқылау, шешімді дәлелдеп қорғау", { text: "Қосымша", options: { color: H.accent5, align: "center" } }],
  ];
  s.addTable(rows, { x: 0.6, y: 1.5, w: 7.9, colW: [2.3, 4.45, 1.15], fontSize: 13, color: H.dk1, border: { type: "solid", pt: 0.75, color: "C9DADF" }, fill: { color: "FFFFFF" }, valign: "middle", rowH: 0.72, margin: 0.08, objectName: "competency-table" });
  // triad on right
  const tri = [[9.55, 1.6, "Пәндік", C.accent1], [11.1, 1.6, "ЖИ сауаттылығы", C.accent5], [10.32, 2.9, "Сыни ойлау", C.accent3]];
  tri.forEach(t => {
    s.addShape(pres.shapes.OVAL, { x: t[0] - 0.75, y: t[1], w: 2.1, h: 2.1, fill: { color: t[3], transparency: 25 }, line: { color: "FFFFFF", width: 1 }, objectName: "triad" });
  });
  txt(s, "Пәндік", { x: 8.85, y: 2.2, w: 1.6, h: 0.4, fontSize: 14, bold: true, color: C.background1, align: "center" });
  txt(s, "ЖИ сауат-тылығы", { x: 10.95, y: 2.1, w: 1.3, h: 0.6, fontSize: 13, bold: true, color: C.background1, align: "center" });
  txt(s, "Сыни ойлау", { x: 9.65, y: 4.25, w: 1.8, h: 0.4, fontSize: 14, bold: true, color: C.background1, align: "center" });
  card(s, 8.8, 5.25, 3.9, 1.55, C.background2, "rec");
  txt(s, [
    { text: "Ұсыныс: ", options: { bold: true, color: C.accent3 } },
    { text: "өзек ретінде сыни ойлау, оған пәндік құзырет пен ЖИ сауаттылығы қызмет етеді. Үшеуі бір тапсырмада қатар дамиды." },
  ], { x: 9.0, y: 5.35, w: 3.55, h: 1.4, fontSize: 13, valign: "middle" });
  s.addNotes("Сұрақ: қай құзыретке тоқталған тиімді? Менің ұсынысым – үштік модель. Өзек – сыни ойлау, себебі диссертацияның тақырыбы соған арналған. Оған пәндік құзырет (Слэйтер ережесі, периодтық заң) мазмұн береді, ал ЖИ сауаттылығы құрал болады. Зерттеушілік пен коммуникативтік құзыреттер қосымша, олар топтық жұмыста өздігінен дамиды. Бір тапсырмада үшеуі қатар жұмыс істейді: студент Z*-ті өзі есептейді (пәндік), ЖИ-ден сұрайды (ЖИ сауаттылығы) және екеуін салыстырып, қатені табады (сыни ойлау). (≈3 мин)");

  // ================= 7. CRITICAL THINKING =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Теориялық негіздер" });
  s.addText("Сыни ойлау: алты негізгі дағды", { placeholder: "title" });
  txt(s, "P. Facione (1990) «Дельфи есебі» бойынша сыни ойлау дағдылары және олардың Слэйтер ережесі тақырыбындағы көрінісі", { x: 0.6, y: 1.25, w: 12.1, h: 0.4, fontSize: 14, color: C.accent6 });
  const ct = [
    [I.eye, "Интерпретациялау", "Иондану энергиясының графигін «оқу», заңдылық пен ауытқуды көру"],
    [I.puzzle, "Талдау", "Электрондарды Слэйтер топтарына бөліп, әр топтың үлесін анықтау"],
    [I.scale, "Бағалау", "ЖИ жауабының және өз есептеуінің дұрыстығын тексеру"],
    [I.bulb, "Қорытынды шығару", "Z* мәнінен атом радиусы мен иондануды болжау"],
    [I.comments, "Түсіндіру", "Неге Fe²⁺ түзілгенде алдымен 4s электрондары кететінін дәлелдеу"],
    [I.sync, "Өзін-өзі реттеу", "Өз қатесін тауып түзету, модельдің шектеулерін мойындау"],
  ];
  ct.forEach((c, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.1, y = 1.85 + row * 2.5;
    card(s, x, y, 3.9, 2.25, C.background2, "ct-card");
    circleIcon(s, c[0], x + 0.25, y + 0.25, 0.7, row === 0 ? C.accent1 : C.accent5);
    txt(s, c[1], { x: x + 1.1, y: y + 0.3, w: 2.7, h: 0.6, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
    txt(s, c[2], { x: x + 0.25, y: y + 1.1, w: 3.45, h: 1.05, fontSize: 14 });
  });
  s.addNotes("Сыни ойлауды өлшенетін етіп анықтау үшін P. Facione-ның «Дельфи есебін» негізге аламын. Мұнда 46 сарапшының келісімімен алты негізгі дағды бөлінген: интерпретациялау, талдау, бағалау, қорытынды шығару, түсіндіру және өзін-өзі реттеу. Әр дағдыны тақырыптағы нақты әрекетпен байланыстырдым. Мысалы, «бағалау» – ЖИ жауабын тексеру, ал «өзін-өзі реттеу» – Слэйтер моделінің шектеулерін мойындау. Кейін бағалау рубрикасы осы дағдыларға сүйенеді. (≈3 мин)");

  // ================= 8. PROBLEM-BASED LEARNING =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Теориялық негіздер" });
  s.addText("Проблемалық оқыту: кезеңдер мен деңгейлер", { placeholder: "title" });
  const st8 = ["Проблемалық жағдаят туғызу", "Мәселені тұжырымдау", "Болжам ұсыну", "Болжамды тексеру", "Қорытынды және рефлексия"];
  const st8d = ["Таңданыс, қайшылық: «Неге олай?»", "Нақты сұрақ қою", "Топта нұсқалар ұсыну", "Есептеу + ЖИ + анықтамалық деректер", "Модельді бағалау, өз ойын қайта қарау"];
  st8.forEach((t, i) => {
    const x = 0.6 + i * 2.47;
    s.addShape(pres.shapes.OVAL, { x: x + 0.7, y: 1.55, w: 0.8, h: 0.8, fill: { color: i === 3 ? C.accent2 : C.accent1 }, line: { color: "FFFFFF" }, objectName: "step" });
    txt(s, String(i + 1), { x: x + 0.7, y: 1.55, w: 0.8, h: 0.8, fontSize: 22, bold: true, color: i === 3 ? C.text1 : C.background1, align: "center", valign: "middle" });
    if (i < 4) s.addShape(pres.shapes.LINE, { x: x + 1.6, y: 1.95, w: 2.27, h: 0, line: { color: C.accent6, width: 1.5, dashType: "dash" }, objectName: "connector" });
    txt(s, t, { x: x, y: 2.5, w: 2.2, h: 0.7, fontSize: 15, bold: true, color: C.text2, align: "center" });
    txt(s, st8d[i], { x: x, y: 3.2, w: 2.2, h: 0.8, fontSize: 13, align: "center" });
  });
  txt(s, "4-кезеңде ЖИ құралдары қолданылады", { x: 7.4, y: 4.05, w: 3.0, h: 0.3, fontSize: 11, italic: true, color: C.accent2, align: "center" });
  // levels
  const lv = [["Проблемалық баяндау", "Оқытушы мәселені қойып, шешу жолын өзі көрсетеді"], ["Ішінара іздену (эвристикалық)", "Студент оқытушы мен ЖИ-дің бағыттаушы сұрақтарымен шешеді"], ["Зерттеушілік", "Студент мәселені дербес шешіп, модельдің шектеулерін зерттейді"]];
  lv.forEach((l, i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 4.6, 3.9, 1.9, i === 1 ? C.accent1 : C.background2, "level");
    txt(s, `${i + 1}-деңгей · ${l[0]}`, { x: x + 0.25, y: 4.75, w: 3.45, h: 0.6, fontSize: 15, bold: true, color: i === 1 ? C.background1 : C.text2 });
    txt(s, l[1], { x: x + 0.25, y: 5.4, w: 3.45, h: 1.0, fontSize: 14, color: i === 1 ? C.background1 : C.text1 });
  });
  txt(s, "Дереккөз: Махмутов М.И. (1975); Лернер И.Я. (1981)", { x: 0.6, y: 6.65, w: 12.1, h: 0.3, fontSize: 11, color: C.accent6 });
  s.addNotes("Проблемалық оқытудың классикалық құрылымын М.И. Махмутов пен И.Я. Лернер еңбектерінен аламыз. Бес кезең бар: проблемалық жағдаят, мәселені тұжырымдау, болжам, тексеру және қорытынды. ЖИ құралдары негізінен 4-кезеңде, болжамды тексеруде қолданылады. Бірақ ЖИ шешімді алмастырмауы керек, сондықтан мен 2-деңгейге, яғни ішінара іздену деңгейіне баса назар аударамын: ЖИ дайын жауап бермей, бағыттаушы сұрақтар қояды. (≈3 мин)");

  // ================= 9. AI IN EDUCATION =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Теориялық негіздер" });
  s.addText("ChatGPT және Claude: рөлдері мен тәуекелдері", { placeholder: "title" });
  const roles = [
    [I.comments, "Сократтық тәлімгер", "Дайын жауап бермей, бағыттаушы сұрақтар қояды"],
    [I.check, "Тексеруші", "Студенттің есептеуін тексеріп, қатені көрсетеді"],
    [I.warn, "«Қате жіберетін серіктес»", "Жауабын студент сынап, түзетеді (Clark, 2023)"],
    [I.laptop, "Оқытушыға көмекші", "Проблемалық жағдаяттар мен тапсырма нұсқаларын әзірлейді"],
  ];
  txt(s, "Мүмкіндіктер (ЖИ-дің дидактикалық рөлдері)", { x: 0.6, y: 1.4, w: 7, h: 0.4, fontSize: 17, bold: true, color: C.accent1 });
  roles.forEach((r, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 3.85, y = 1.95 + row * 2.2;
    card(s, x, y, 3.65, 2.0, C.background2, "role");
    circleIcon(s, r[0], x + 0.2, y + 0.2, 0.65, C.accent1);
    txt(s, r[1], { x: x + 1.0, y: y + 0.2, w: 2.55, h: 0.65, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    txt(s, r[2], { x: x + 0.2, y: y + 1.0, w: 3.3, h: 0.9, fontSize: 14 });
  });
  card(s, 8.5, 1.4, 4.2, 5.35, C.text2, "risk");
  txt(s, "Тәуекелдер және шешімдер", { x: 8.8, y: 1.6, w: 3.7, h: 0.4, fontSize: 17, bold: true, color: C.accent2 });
  txt(s, [
    { text: "«Галлюцинация»: қате сандар мен ойдан шығарылған дереккөздер", options: { bullet: true, breakLine: true } },
    { text: "→ анықтамалық деректермен міндетті салыстыру", options: { color: H.accent2, breakLine: true } },
    { text: "Ойлауды ЖИ-ге тапсыру, тәуелділік", options: { bullet: true, breakLine: true } },
    { text: "→ алдымен өз болжамы, содан кейін ЖИ", options: { color: H.accent2, breakLine: true } },
    { text: "Академиялық адалдық", options: { bullet: true, breakLine: true } },
    { text: "→ ЖИ-мен диалогты есепке қоса тапсыру", options: { color: H.accent2 } },
  ], { x: 8.8, y: 2.15, w: 3.7, h: 4.0, fontSize: 14, color: C.background1, paraSpaceAfter: 6 });
  txt(s, "UNESCO (2023); Kasneci et al. (2023)", { x: 8.8, y: 6.3, w: 3.7, h: 0.3, fontSize: 11, color: C.background2 });
  s.addNotes("Зерттеуде ChatGPT пен Claude қолданылады. ЖИ-ге төрт дидактикалық рөл беремін. Ең маңыздысы – үшінші рөл, «қате жіберетін серіктес»: Clark 2023 жылы ChatGPT жауаптары жиі қате болғанымен, оларды сынап, жақсарту тапсырмаларына пайдалануға болатынын жазады. Тәуекелдер де бар: галлюцинация, тәуелділік, академиялық адалдық. ЮНЕСКО-ның 2023 жылғы нұсқаулығы мен Kasneci және әріптестерінің жұмысы осы тәуекелдерді басқарудың жолдарын ұсынады. Мен үш ережені қолданамын: анықтамалық деректермен салыстыру, алдымен өз болжамын жасау және ЖИ-мен диалогты есепке қоса тапсыру. (≈3 мин)");

  // ================= 10. INTEGRATION MODEL =================
  pres.addSection({ title: "Кіріктіру моделі" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріктіру моделі" });
  s.addText("Кіріктіру моделі", { placeholder: "title" });
  const blocks = [
    ["МАҚСАТ", "Сыни ойлау құзыреттілігі (Facione)", C.accent3, I.target],
    ["МАЗМҰН", "Атом құрылысы · Периодтық заң · Слэйтер ережесі", C.accent1, I.atom],
    ["ӘДІС", "Проблемалық оқыту (5 кезең, 3 деңгей)", C.accent5, I.puzzle],
    ["ҚҰРАЛ", "ChatGPT, Claude: тәлімгер, тексеруші, сынға алынатын серіктес", C.accent4, I.robot],
    ["НӘТИЖЕ", "Рубрика, сыни ойлау тесті, ЖИ-диалог портфолиосы", C.accent2, I.flag],
  ];
  blocks.forEach((b, i) => {
    const x = 0.6 + i * 2.475;
    card(s, x, 1.55, 2.2, 3.15, b[2], "model-block");
    circleIcon(s, b[3], x + 0.7, 1.75, 0.8, C.text2);
    txt(s, b[0], { x: x + 0.12, y: 2.7, w: 1.96, h: 0.4, fontSize: 16, bold: true, color: i === 4 ? C.text1 : C.background1, align: "center" });
    txt(s, b[1], { x: x + 0.12, y: 3.15, w: 1.96, h: 1.45, fontSize: 13, color: i === 4 ? C.text1 : C.background1, align: "center" });
    if (i < 4) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 2.24, y: 2.95, w: 0.2, h: 0.35, fill: { color: C.accent6 }, line: { color: C.accent6 }, objectName: "arrow" });
  });
  card(s, 0.6, 4.95, 12.1, 1.85, C.background2, "why-topic");
  circleIcon(s, I.atom, 0.85, 5.3, 1.0, C.accent1);
  txt(s, [
    { text: "Неге «Атом құрылысы және Периодтық заң»? ", options: { bold: true, color: C.text2, breakLine: true } },
    { text: "Бұл бейорганикалық химияның іргетасы: элементтердің барлық қасиеті осыдан шығады. Слэйтер ережесі «электрондық конфигурация» мен «периодтық қасиеттер» арасындағы көпір: қарапайым есептеу арқылы заңдылықты түсіндіреді, ал оның шектеулері проблемалық жағдаят туғызуға таптырмас материал." },
  ], { x: 2.1, y: 5.1, w: 10.4, h: 1.6, fontSize: 15, valign: "middle" });
  s.addNotes("Бұл слайд – баяндаманың өзегі. Модель бес компоненттен тұрады. Мақсат – сыни ойлау. Мазмұн – Атом құрылысы, Периодтық заң және Слэйтер ережесі. Әдіс – проблемалық оқыту. Құрал – ChatGPT пен Claude. Нәтиже рубрика, тест және ЖИ-мен диалог портфолиосы арқылы бағаланады. Неге осы тақырып? Себебі Слэйтер ережесі көпір рөлін атқарады: ол электрондық конфигурацияны периодтық қасиеттермен сандық түрде байланыстырады. Әрі ол – жуықтау модель, сондықтан шектеулері бар. Модельдің шектеуі – сыни ойлау үшін ең жақсы материал. (≈3 мин)");

  // ================= 11. SLATER RULES =================
  pres.addSection({ title: "Слэйтер ережесі" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Слэйтер ережесі" });
  s.addText("Слэйтер ережесі: ядроның тиімді заряды", { placeholder: "title" });
  card(s, 0.6, 1.5, 4.6, 2.2, C.text2, "formula");
  txt(s, "Z* = Z − S", { x: 0.6, y: 1.65, w: 4.6, h: 1.0, fontSize: 44, bold: true, color: C.background1, align: "center", valign: "middle", fontFace: THEME.headFontFace });
  txt(s, [
    { text: "Z* – ядроның тиімді заряды", options: { breakLine: true } },
    { text: "Z – ядро заряды (реттік нөмір)", options: { breakLine: true } },
    { text: "S – экрандау (бүркеу) тұрақтысы", options: {} },
  ], { x: 0.9, y: 2.65, w: 4.1, h: 1.0, fontSize: 14, color: C.background2, align: "center" });
  txt(s, "Электрондарды топтарға жазу:", { x: 0.6, y: 3.95, w: 4.6, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
  const grp = ["1s", "2s,2p", "3s,3p", "3d", "4s,4p", "4d", "4f", "5s,5p"];
  grp.forEach((g, i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = 0.6 + col * 1.17, y = 4.4 + row * 0.75;
    card(s, x, y, 1.07, 0.6, C.background2, "group");
    txt(s, `(${g})`, { x, y, w: 1.07, h: 0.6, fontSize: 13, bold: true, color: C.accent1, align: "center", valign: "middle" });
  });
  txt(s, "Слэйтер, 1930", { x: 0.6, y: 6.0, w: 4.6, h: 0.3, fontSize: 11, color: C.accent6 });
  const sr = [
    [{ text: "Экрандаушы электрондар", options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 } } }, { text: "ns, np электрон үшін", options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 }, align: "center" } }, { text: "nd, nf электрон үшін", options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 }, align: "center" } }],
    ["Оң жақтағы топтар", { text: "0", options: { align: "center" } }, { text: "0", options: { align: "center" } }],
    ["Сол топтағы басқа электрондар", { text: "0,35 (1s үшін 0,30)", options: { align: "center" } }, { text: "0,35", options: { align: "center" } }],
    ["(n − 1) қабат", { text: "0,85", options: { align: "center", bold: true, color: H.accent1 } }, { text: "1,00", options: { align: "center" } }],
    ["(n − 2) және одан төмен", { text: "1,00", options: { align: "center" } }, { text: "1,00", options: { align: "center" } }],
  ];
  s.addTable(sr, { x: 5.6, y: 1.5, w: 7.1, colW: [2.9, 2.2, 2.0], fontSize: 14, color: H.dk1, border: { type: "solid", pt: 0.75, color: "C9DADF" }, fill: { color: "FFFFFF" }, valign: "middle", rowH: 0.62, margin: 0.08, objectName: "slater-table" });
  card(s, 5.6, 4.85, 7.1, 1.5, C.background2, "meaning");
  circleIcon(s, I.bulb, 5.85, 5.15, 0.85, C.accent2);
  txt(s, "Ішкі электрондар ядро зарядын «бүркейді». Сыртқы электрон ядроның толық зарядын емес, тек Z* бөлігін «сезеді». Z* неғұрлым үлкен болса, электрон соғұрлым күшті тартылады.", { x: 6.9, y: 4.95, w: 5.6, h: 1.3, fontSize: 14, valign: "middle" });
  s.addNotes("Слэйтер ережесін 1930 жылы Джон Слэйтер ұсынған. Негізгі формула: Z* = Z − S. Электрондарды арнайы топтарға жазамыз: (1s)(2s,2p)(3s,3p)(3d)(4s,4p) және т.б. Есептелетін электроннан оң жақтағы электрондар экрандамайды. Сол топтағы әр электрон 0,35 береді, (n−1) қабат 0,85 береді, ал одан да ішкі қабаттар 1,00 береді. d және f электрондар үшін сол жақтағы барлық электрондар 1,00-ден береді. Физикалық мағынасы: ішкі электрондар ядроны «бүркейді». (≈3 мин)");

  // ================= 12. WORKED EXAMPLE =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Слэйтер ережесі" });
  s.addText("Есептеу мысалы: хлордың 3p электроны", { placeholder: "title" });
  txt(s, "Cl (Z = 17):   (1s²) (2s² 2p⁶) (3s² 3p⁵)", { x: 0.6, y: 1.4, w: 7.4, h: 0.5, fontSize: 20, bold: true, color: C.text2 });
  const steps = [
    ["(1s²)", "2 × 1,00", "= 2,00", "(n − 2) қабат", C.accent5],
    ["(2s² 2p⁶)", "8 × 0,85", "= 6,80", "(n − 1) қабат", C.accent1],
    ["(3s² 3p⁵)", "6 × 0,35", "= 2,10", "сол топ (17-нің өзінен басқа 6 электрон)", C.accent2],
  ];
  steps.forEach((st, i) => {
    const y = 2.1 + i * 0.95;
    card(s, 0.6, y, 1.9, 0.75, st[4], "grp");
    txt(s, st[0], { x: 0.6, y, w: 1.9, h: 0.75, fontSize: 16, bold: true, color: i === 2 ? C.text1 : C.background1, align: "center", valign: "middle" });
    txt(s, st[1], { x: 2.7, y, w: 1.4, h: 0.75, fontSize: 18, valign: "middle" });
    txt(s, st[2], { x: 4.1, y, w: 1.1, h: 0.75, fontSize: 18, bold: true, valign: "middle" });
    txt(s, st[3], { x: 5.3, y, w: 2.7, h: 0.75, fontSize: 13, color: C.accent6, valign: "middle" });
  });
  card(s, 0.6, 5.05, 7.4, 1.6, C.text2, "result");
  txt(s, [
    { text: "S = 2,00 + 6,80 + 2,10 = 10,90", options: { breakLine: true } },
    { text: "Z* = 17 − 10,90 = 6,10", options: { bold: true, color: H.accent2, fontSize: 26 } },
  ], { x: 0.9, y: 5.15, w: 6.9, h: 1.4, fontSize: 18, color: C.background1, valign: "middle" });
  // comparison table
  txt(s, "Слэйтер ережесі мен Клементи–Раймонди мәндері", { x: 8.4, y: 1.4, w: 4.3, h: 0.6, fontSize: 15, bold: true, color: C.text2 });
  const cmp = [
    [{ text: "Электрон", options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 } } }, { text: "Слэйтер", options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 }, align: "center" } }, { text: "К–Р (1963)", options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 }, align: "center" } }],
    ["Li 2s", { text: "1,30", options: { align: "center" } }, { text: "1,28", options: { align: "center" } }],
    ["F 2p", { text: "5,20", options: { align: "center" } }, { text: "5,10", options: { align: "center" } }],
    ["Na 3s", { text: "2,20", options: { align: "center" } }, { text: "2,51", options: { align: "center" } }],
    ["Cl 3p", { text: "6,10", options: { align: "center" } }, { text: "6,12", options: { align: "center" } }],
    [{ text: "K 4s", options: { bold: true, color: H.accent3 } }, { text: "2,20", options: { align: "center", bold: true, color: H.accent3 } }, { text: "3,50", options: { align: "center", bold: true, color: H.accent3 } }],
  ];
  s.addTable(cmp, { x: 8.4, y: 2.1, w: 4.3, colW: [1.5, 1.4, 1.4], fontSize: 14, color: H.dk1, border: { type: "solid", pt: 0.75, color: "C9DADF" }, fill: { color: "FFFFFF" }, valign: "middle", rowH: 0.48, margin: 0.06, objectName: "cmp-table" });
  card(s, 8.4, 5.2, 4.3, 1.45, C.background2, "q");
  txt(s, [
    { text: "Сыни сұрақ: ", options: { bold: true, color: C.accent3 } },
    { text: "калий үшін айырмашылық неге үлкен? Модельдің қай жорамалы «жұмыс істемейді»?" },
  ], { x: 8.6, y: 5.3, w: 3.95, h: 1.25, fontSize: 14, valign: "middle" });
  s.addNotes("Хлордың 3p электроны үшін есептейміз. Ішкі 1s қабатындағы 2 электрон толық бүркейді: 2×1,00. Екінші қабаттың 8 электроны (n−1) қабат: 8×0,85 = 6,80. Үшінші қабаттағы қалған 6 электрон: 6×0,35 = 2,10. Барлығы S = 10,90, Z* = 6,10. Оң жақтағы кестеде Клементи мен Раймондидің 1963 жылғы, өздігінен келісілген өріс (ӨКӨ) әдісімен есептелген мәндері берілген. Хлор үшін сәйкестік өте жақсы, ал калийде айырмашылық үлкен. Мұнда студентке сыни сұрақ қоямыз: модель неге калийде қателеседі? Жауап: Слэйтер ережесі 4s электронының ішкі қабаттарға енуін (пенетрациясын) ескермейді. (≈3 мин)");

  // ================= 13. PERIODICITY CHARTS =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Слэйтер ережесі" });
  s.addText("Z* және периодтық заңдылықтар", { placeholder: "title" });
  const cats = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"];
  s.addChart(pres.charts.LINE, [
    { name: "2-период (Li–Ne)", labels: cats, values: [1.30, 1.95, 2.60, 3.25, 3.90, 4.55, 5.20, 5.85] },
    { name: "3-период (Na–Ar)", labels: cats, values: [2.20, 2.85, 3.50, 4.15, 4.80, 5.45, 6.10, 6.75] },
  ], Object.assign(chartBase(), {
    x: 0.6, y: 1.4, w: 6.0, h: 4.2, title: "Сыртқы электронның Z* мәні (Слэйтер)", chartColors: [H.accent1, H.accent5],
    lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 7, showLegend: true, legendPos: "b",
    valAxisMinVal: 0, valAxisMaxVal: 8, valAxisMajorUnit: 2, showCatAxisTitle: true, catAxisTitle: "Топ", catAxisTitleFontFace: "+mn-lt", catAxisTitleFontSize: 12, catAxisTitleColor: "1B2430",
  }));
  s.addChart(pres.charts.LINE, [
    { name: "2-период (Li–Ne)", labels: cats, values: [520, 899, 801, 1086, 1402, 1314, 1681, 2081] },
    { name: "3-период (Na–Ar)", labels: cats, values: [496, 738, 578, 787, 1012, 1000, 1251, 1521] },
  ], Object.assign(chartBase(), {
    x: 6.75, y: 1.4, w: 6.0, h: 4.2, title: "Бірінші иондану энергиясы, кДж/моль", chartColors: [H.accent1, H.accent5],
    lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 7, showLegend: true, legendPos: "b",
    valAxisMinVal: 0, valAxisMaxVal: 2400, valAxisMajorUnit: 600, showCatAxisTitle: true, catAxisTitle: "Топ", catAxisTitleFontFace: "+mn-lt", catAxisTitleFontSize: 12, catAxisTitleColor: "1B2430",
  }));
  card(s, 0.6, 5.8, 6.0, 1.05, C.background2, "c1");
  txt(s, [{ text: "Заңдылық: ", options: { bold: true, color: C.accent1 } }, { text: "период бойында Z* біркелкі (+0,65) өседі, сондықтан иондану энергиясы артып, радиус кішірейеді." }], { x: 0.8, y: 5.85, w: 5.7, h: 0.95, fontSize: 14, valign: "middle" });
  card(s, 6.75, 5.8, 5.95, 1.05, C.background2, "c2");
  txt(s, [{ text: "Проблема: ", options: { bold: true, color: C.accent3 } }, { text: "Z* біркелкі өседі, ал неге Be > B және N > O? Слэйтер моделі бұл «секірістерді» түсіндіре ала ма?" }], { x: 6.95, y: 5.85, w: 5.65, h: 0.95, fontSize: 14, valign: "middle" });
  s.addNotes("Сол жақ графикте 2 және 3-периодтағы сыртқы электронның Z* мәні көрсетілген: Слэйтер бойынша ол әр элемент сайын 0,65-ке біркелкі өседі. Бұл период бойында иондану энергиясының өсуін және радиустың кішіреюін жақсы түсіндіреді. Бірақ оң жақ графикке қараңыздар: иондану энергиясы бірқалыпты өспейді. Be-ден B-ға және N-нен O-ға өткенде энергия төмендейді. Студенттер өздері қайшылықты байқауы керек: модель біркелкі өсуді болжайды, ал тәжірибелік деректер басқаша. Жауабы: Слэйтер ережесі s және p электрондарын бір топқа біріктіреді, ал электрондардың жұптасу энергиясын ескермейді. Модельдің шектеулерін тану – сыни ойлаудың нақ өзі. Иондану энергиясы кДж/моль бірлігімен, анықтамалық мәндер. (≈3,5 мин)");

  // ================= 14. PROBLEM 1: Fe =================
  pres.addSection({ title: "Проблемалық жағдаяттар" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Проблемалық жағдаяттар" });
  s.addText("Проблемалық жағдаят 1: темір иондары", { placeholder: "title" });
  card(s, 0.6, 1.45, 12.1, 1.25, C.text2, "problem");
  circleIcon(s, I.question, 0.85, 1.65, 0.85, C.accent2);
  txt(s, "Fe: [Ar] 3d⁶ 4s². Клечковский ережесі бойынша 4s орбиталь 3d-дан бұрын толады. Онда неге Fe → Fe²⁺ кезінде алдымен 4s электрондары кетеді? Соңғы келген электрон бірінші кетуі керек емес пе?", { x: 1.95, y: 1.5, w: 10.55, h: 1.15, fontSize: 16, color: C.background1, valign: "middle" });
  // two calc cards
  const fe = [
    ["4s электроны", "(1s²)(2s²2p⁶)(3s²3p⁶)(3d⁶)(4s²)", "S = 1×0,35 + 14×0,85 + 10×1,00 = 22,25", "Z* = 26 − 22,25 = 3,75", 3.75, C.accent5],
    ["3d электроны", "(1s²)(2s²2p⁶)(3s²3p⁶)(3d⁶)", "S = 5×0,35 + 18×1,00 = 19,75", "Z* = 26 − 19,75 = 6,25", 6.25, C.accent1],
  ];
  fe.forEach((f, i) => {
    const y = 2.95 + i * 1.75;
    card(s, 0.6, y, 7.3, 1.55, C.background2, "fe-card");
    txt(s, f[0], { x: 0.85, y: y + 0.12, w: 3, h: 0.35, fontSize: 16, bold: true, color: f[5] });
    txt(s, f[2], { x: 0.85, y: y + 0.52, w: 6.9, h: 0.4, fontSize: 15 });
    txt(s, f[3], { x: 0.85, y: y + 0.95, w: 6.9, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    // bar
    const bw = (f[4] / 7) * 3.9;
    s.addShape(pres.shapes.RECTANGLE, { x: 8.3, y: y + 0.45, w: bw, h: 0.65, fill: { color: f[5] }, line: { color: f[5] }, objectName: "zbar" });
    txt(s, f[4].toFixed(2).replace(".", ","), { x: 8.3 + bw + 0.1, y: y + 0.45, w: 0.9, h: 0.65, fontSize: 18, bold: true, color: f[5], valign: "middle" });
  });
  txt(s, "Z* салыстыру", { x: 8.3, y: 2.95, w: 3, h: 0.35, fontSize: 13, color: C.accent6 });
  txt(s, [
    { text: "Қорытынды: ", options: { bold: true, color: C.accent1 } },
    { text: "4s электронын ядро әлсіз тартады (Z* = 3,75 < 6,25), сондықтан ол бірінші кетеді. Толу реті ≠ иондану реті." },
  ], { x: 0.6, y: 6.5, w: 12.1, h: 0.45, fontSize: 15 });
  s.addNotes("Бірінші проблемалық жағдаят студенттерде нақты қайшылық туғызады. Олар Клечковский ережесін біледі: 4s бірінші толады. Сондықтан логика бойынша 3d электрондары бірінші кетуі керек сияқты. Бірақ тәжірибеде Fe²⁺ конфигурациясы [Ar]3d⁶. Студенттер топта болжам айтып, кейін Слэйтер бойынша екі электрон үшін Z*-ті есептейді: 4s үшін 3,75, 3d үшін 6,25. Демек, 4s электроны әлсіз байланысқан. Осыдан кейін ЖИ-ден сұрайды да, оның түсіндірмесін өз есептеуімен салыстырады. Нәтижесінде студент «толу реті иондану ретімен бірдей емес» деген маңызды қорытындыға өзі келеді. (≈3 мин)");

  // ================= 15. PROBLEM 2: Group 1 =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Проблемалық жағдаяттар" });
  s.addText("Проблемалық жағдаят 2: сілтілік металдар", { placeholder: "title" });
  s.addChart(pres.charts.BAR, [
    { name: "Металдық радиус, пм", labels: ["Li", "Na", "K", "Rb", "Cs"], values: [152, 186, 227, 248, 265] },
  ], Object.assign(chartBase(), {
    x: 0.6, y: 1.4, w: 6.6, h: 4.0, barDir: "col", title: "Атом радиусы, пм", chartColors: [H.accent1], showLegend: false,
    showValue: true, dataLabelPosition: "outEnd", valAxisMinVal: 0, valAxisMaxVal: 300, valAxisMajorUnit: 100, barGapWidthPct: 60,
  }));
  txt(s, "Сыртқы ns электронның Z* мәні (Слэйтер):", { x: 0.6, y: 5.5, w: 6.6, h: 0.35, fontSize: 14, bold: true, color: C.text2 });
  [["Li", "1,30"], ["Na", "2,20"], ["K", "2,20"], ["Rb", "2,20"], ["Cs", "2,20"]].forEach((e, i) => {
    const x = 0.95 + i * 1.25;
    card(s, x, 5.95, 1.05, 0.85, i === 0 ? C.background2 : C.accent2, "zchip");
    txt(s, [{ text: e[0], options: { bold: true, breakLine: true, fontSize: 13 } }, { text: e[1], options: { fontSize: 16, bold: true } }], { x, y: 5.95, w: 1.05, h: 0.85, align: "center", valign: "middle", color: C.text1 });
  });
  card(s, 7.6, 1.4, 5.1, 2.35, C.text2, "p2");
  txt(s, [
    { text: "Қайшылық", options: { bold: true, color: H.accent2, breakLine: true } },
    { text: "Na-дан Cs-ке дейін Z* өзгермейді (2,20), ал ядро заряды 11-ден 55-ке дейін өседі. Онда неге радиус 186-дан 265 пм-ге дейін ұлғаяды және иондану энергиясы азаяды?" },
  ], { x: 7.85, y: 1.55, w: 4.65, h: 2.1, fontSize: 15, color: C.background1 });
  card(s, 7.6, 3.95, 5.1, 2.85, C.background2, "hint");
  txt(s, "Студенттердің ізденіс бағыттары", { x: 7.85, y: 4.1, w: 4.6, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
  txt(s, [
    { text: "Радиусты тек Z* емес, бас кванттық сан n да анықтайды", options: { bullet: true, breakLine: true } },
    { text: "Слэйтер бойынша r ∝ n*² / Z* (n* – тиімді кванттық сан)", options: { bullet: true, breakLine: true } },
    { text: "ЖИ-ден сұрау және оның жауабын Клементи–Раймонди мәндерімен салыстыру", options: { bullet: true } },
  ], { x: 7.85, y: 4.5, w: 4.65, h: 2.2, fontSize: 14, paraSpaceAfter: 6 });
  s.addNotes("Екінші жағдаят 1-топқа арналған. Слэйтер бойынша Na, K, Rb, Cs үшін Z* бірдей, 2,20, ал ядро заряды бес есе өседі. Студенттер: «Z* өзгермесе, неге радиус өседі?» деп ойланады. Шешімі: радиус тек Z*-ке емес, бас кванттық санға да байланысты. Слэйтер ұсынған тиімді кванттық сан n* (1, 2, 3, 3,7, 4,0, 4,2) арқылы r ∝ n*²/Z* қатынасы Li-ден Cs-ке қарай өседі. Бұл жерде студент екі факторды бірге талдауды үйренеді. Сонымен қатар Клементи–Раймонди мәндері бойынша Z* шын мәнінде топ бойымен аздап өседі, ендеше «Z* тұрақты» деген тұжырым да модельдің жуықтауы. Радиустар – металдық радиустың анықтамалық мәндері. (≈3 мин)");

  // ================= 16. CRITIQUE AI =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Проблемалық жағдаяттар" });
  s.addText("Тапсырма: «ЖИ жауабын тексер»", { placeholder: "title" });
  // chat mock
  card(s, 0.6, 1.45, 6.6, 5.35, C.background2, "chat");
  txt(s, "ЖИ жауабының үлгісі (оқытушы әдейі қате енгізген)", { x: 0.85, y: 1.6, w: 6.1, h: 0.35, fontSize: 13, italic: true, color: C.accent6 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.4, y: 2.05, w: 4.6, h: 0.8, fill: { color: C.accent1 }, line: { color: C.accent1 }, rectRadius: 0.1, objectName: "bubble-user" });
  txt(s, "Калийдің 4s электроны үшін Z*-ті Слэйтер ережесімен есепте.", { x: 2.55, y: 2.1, w: 4.3, h: 0.7, fontSize: 14, color: C.background1, valign: "middle" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.85, y: 3.05, w: 5.4, h: 2.15, fill: { color: "FFFFFF" }, line: { color: "C9DADF" }, rectRadius: 0.1, objectName: "bubble-ai" });
  circleIcon(s, I.robot, 0.95, 3.15, 0.5, C.accent5);
  txt(s, [
    { text: "K: (1s²)(2s²2p⁶)(3s²3p⁶)(4s¹)", options: { breakLine: true } },
    { text: "S = 8 × 0,35 + 8 × 0,85 + 2 × 1,00 = 11,6", options: { breakLine: true, color: H.accent3, bold: true } },
    { text: "Z* = 19 − 11,6 = 7,4", options: { breakLine: true, color: H.accent3, bold: true } },
    { text: "Сондықтан калийдің сыртқы электроны өте күшті тартылады, иондану энергиясы жоғары." },
  ], { x: 1.6, y: 3.2, w: 4.5, h: 2.3, fontSize: 14, paraSpaceAfter: 4 });
  txt(s, "Дұрыс жауабы: S = 8×0,85 + 10×1,00 = 16,8; Z* = 2,20", { x: 0.85, y: 5.45, w: 6.1, h: 0.75, fontSize: 15, bold: true, color: C.accent6, valign: "middle" });
  // tasks
  txt(s, "Студенттің әрекеттері", { x: 7.6, y: 1.45, w: 5.1, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
  const acts = [
    [I.search, "Қатені табу", "Қай электрон қай топқа дұрыс жатқызылмаған?"],
    [I.check, "Түзету", "Дұрыс есептеп, Z* = 2,20 мәнін алу"],
    [I.scale, "Шындықпен салыстыру", "Анықтама: IE(K) = 419 кДж/моль, яғни төмен. ЖИ қорытындысы қисынсыз"],
    [I.comments, "Кері сұрау", "ЖИ-ге қатесін көрсетіп, оның реакциясын талдау"],
  ];
  acts.forEach((a, i) => {
    const y = 2.0 + i * 1.2;
    circleIcon(s, a[0], 7.6, y, 0.7, i === 2 ? C.accent3 : C.accent1);
    txt(s, a[1], { x: 8.5, y: y - 0.02, w: 4.2, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    txt(s, a[2], { x: 8.5, y: y + 0.33, w: 4.2, h: 0.75, fontSize: 13 });
  });
  s.addNotes("Бұл тапсырма Clark (2023) ұсынысына негізделген: студенттер ЖИ жауабын сынайды. Мұндағы жауап нақты ЖИ шығарған мәтін емес, оқытушы ЖИ-ге тән қателерді әдейі енгізіп дайындаған үлгі. Қате мынада: 3s3p электрондары (n−1) қабатқа жатады және 0,85 береді, ал ішкі 10 электрон 1,00 береді, ал үлгіде топтар араласып кеткен. Студент қатені тауып, түзетеді, ең бастысы, қорытындыны шындықпен салыстырады: калийдің иондану энергиясы 419 кДж/моль, бұл өте төмен мән. Демек, «күшті тартылады» деген тұжырым есептің қате екенін өзі көрсетеді. Нақты сабақта студенттер ChatGPT пен Claude-тан бір сұрақты сұрап, жауаптарын салыстырады. (≈3 мин)");

  // ================= 17. PROMPTS + LESSON =================
  pres.addSection({ title: "Сабақ үлгісі" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Сабақ үлгісі" });
  s.addText("Сабақ құрылымы (50 минут)", { placeholder: "title" });
  const lesson = [
    [5, "Жағдаят", "Fe²⁺ немесе IE «секірісі»", C.accent3],
    [7, "Болжам", "Топта нұсқалар, ЖИ-сіз", C.accent5],
    [13, "Слэйтер есептеуі", "Z* есептеу, кестемен салыстыру", C.accent1],
    [12, "ЖИ-мен диалог", "ChatGPT / Claude жауабын тексеру", C.accent4],
    [8, "Дәлелдеу", "Топ шешімін қорғау", C.accent5],
    [5, "Рефлексия", "Модель шегі, өзін бағалау", C.accent2],
  ];
  let xx = 0.6; const totalW = 12.1;
  lesson.forEach((l, i) => {
    const w = (l[0] / 50) * totalW;
    s.addShape(pres.shapes.RECTANGLE, { x: xx, y: 1.5, w: w - 0.05, h: 0.75, fill: { color: l[3] }, line: { color: l[3] }, objectName: "timeline" });
    txt(s, `${l[0]} мин`, { x: xx, y: 1.5, w: w - 0.05, h: 0.75, fontSize: 14, bold: true, color: i === 5 ? C.text1 : C.background1, align: "center", valign: "middle" });
    txt(s, l[1], { x: xx, y: 2.35, w: w - 0.1, h: 0.6, fontSize: 13, bold: true, color: C.text2, align: "center" });
    txt(s, l[2], { x: xx + 0.03, y: 2.95, w: w - 0.16, h: 0.85, fontSize: 11, color: C.text1, align: "center" });
    xx += w;
  });
  txt(s, "Промпт (сұраныс) үлгілері", { x: 0.6, y: 3.95, w: 8, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
  const prompts = [
    ["Сократтық тәлімгер", "«Сен химия оқытушысысың. Маған дайын жауап берме. Слэйтер ережесі бойынша Cl атомының 3p электроны үшін Z*-ті есептеуге жетелейтін сұрақтарды бір-бірлеп қой.»"],
    ["Тексеруші", "«Менің есептеуім: S = 2×1,00 + 8×0,85 + 6×0,35. Қай қадам дұрыс, қайсысы қате? Тек себебін түсіндір, түзетуді өзім жасаймын.»"],
    ["Модельді сынау", "«Слэйтер ережесі B мен O-ның иондану энергиясындағы ауытқуды түсіндіре ала ма? Модельдің шектеулерін атап, дереккөз көрсет.»"],
  ];
  prompts.forEach((p, i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 4.45, 3.9, 2.35, C.background2, "prompt");
    txt(s, p[0], { x: x + 0.2, y: 4.55, w: 3.5, h: 0.35, fontSize: 14, bold: true, color: C.accent1 });
    txt(s, p[1], { x: x + 0.2, y: 4.95, w: 3.5, h: 1.8, fontSize: 14, italic: true });
  });
  s.addNotes("Бір академиялық сағат, яғни 50 минуттық сабақтың құрылымы. Маңызды қағида: алғашқы 25 минутта ЖИ қолданылмайды, студент алдымен өзі ойлап, өзі есептейді. ЖИ тек 4-кезеңде, тексеру үшін қосылады. Төменде үш промпт үлгісі берілген. Біріншісі ЖИ-ді Сократ әдісімен сұрақ қоятын тәлімгерге айналдырады, екіншісі тек қатені көрсетеді, ал үшіншісі модельдің шектеулерін талқылауға бағытталған. Студенттер ChatGPT пен Claude жауаптарын салыстырып, айырмашылықтарды талдайды. (≈3 мин)");

  // ================= 18. RUBRIC =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Сабақ үлгісі" });
  s.addText("Сыни ойлауды бағалау рубрикасы", { placeholder: "title" });
  const hd = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 }, align: "center" } });
  const rub = [
    [hd("Критерий"), hd("1 – Бастапқы"), hd("2 – Жеткілікті"), hd("3 – Жоғары")],
    [{ text: "Талдау", options: { bold: true } }, "Электрондарды топтарға қатемен бөледі", "Дұрыс бөледі, бірақ коэффициентте қателеседі", "Z*-ті дұрыс есептеп, әр топтың үлесін түсіндіреді"],
    [{ text: "ЖИ жауабын бағалау", options: { bold: true } }, "ЖИ жауабын тексерусіз қабылдайды", "Қатені табады, себебін толық түсіндірмейді", "Қатені тауып, анықтамалық деректермен дәлелдейді"],
    [{ text: "Қорытынды шығару", options: { bold: true } }, "Z* пен қасиет арасында байланыс жоқ", "Заңдылықты сипаттайды", "Z* арқылы қасиеттерді болжап, ерекшеліктерді түсіндіреді"],
    [{ text: "Модель шектеулері", options: { bold: true } }, "Модельді абсолютті шындық деп санайды", "Ауытқуды байқайды", "Ауытқудың себебін және модельдің қолданылу шегін атайды"],
    [{ text: "Рефлексия", options: { bold: true } }, "Өз қатесін көрмейді", "Оқытушының көмегімен түзетеді", "Өз ойлау жолын дербес талдап, түзетеді"],
  ];
  s.addTable(rub, { x: 0.6, y: 1.45, w: 8.6, colW: [1.9, 2.2, 2.25, 2.25], fontSize: 12, color: H.dk1, border: { type: "solid", pt: 0.75, color: "C9DADF" }, fill: { color: "FFFFFF" }, valign: "middle", rowH: [0.5, 0.86, 0.86, 0.86, 0.86, 0.86], margin: 0.07, objectName: "rubric" });
  card(s, 9.5, 1.45, 3.2, 5.3, C.background2, "measure");
  txt(s, "Өлшеу құралдары", { x: 9.75, y: 1.6, w: 2.8, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
  [[I.check, "Рубрика", "5 критерий × 3 деңгей, барлығы 15 балл"], [I.book, "Сыни ойлау тесті", "Эксперимент басында және соңында"], [I.laptop, "ЖИ-диалог портфолиосы", "Сұраныстар мен тексерулер тарихы"]].forEach((m, i) => {
    const y = 2.15 + i * 1.5;
    circleIcon(s, m[0], 9.75, y, 0.6, C.accent1);
    txt(s, m[1], { x: 10.5, y: y - 0.02, w: 2.1, h: 0.6, fontSize: 14, bold: true, color: C.text2 });
    txt(s, m[2], { x: 10.5, y: y + 0.62, w: 2.1, h: 0.8, fontSize: 12 });
  });
  s.addNotes("Сыни ойлау дамығанын қалай білеміз? Бес критерийлі рубрика ұсынамын. Ол Facione дағдыларына және Блумның қайта қаралған таксономиясының (Anderson, Krathwohl, 2001) жоғары деңгейлеріне, яғни талдау, бағалау және жасауға сүйенеді. Әр критерий үш деңгейде бағаланады, барлығы 15 балл. Рубрикадан басқа эксперимент басында және соңында сыни ойлау тесті алынады, ал ЖИ-мен диалог портфолиосы студенттің ЖИ-ді қалай қолданғанын көрсетеді. Бұл үш құрал педагогикалық эксперименттің деректерін береді. (≈3 мин)");

  // ================= 19. CONCLUSION (dark) =================
  pres.addSection({ title: "Қорытынды" });
  s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Қорытынды" });
  txt(s, "Қорытынды", { x: 0.7, y: 0.6, w: 8, h: 0.8, fontSize: 36, bold: true, color: C.background1, fontFace: THEME.headFontFace });
  const concl = [
    [I.layers, "Құзыреттілік тәсіл", "сыни ойлауды өзек етіп, пәндік құзырет пен ЖИ сауаттылығымен біріктіруге мүмкіндік береді"],
    [I.atom, "Слэйтер ережесі", "Атом құрылысы мен Периодтық заңды сандық түрде байланыстырады. Оның шектеулері де проблемалық жағдаятқа арқау болады"],
    [I.robot, "ChatGPT мен Claude", "дайын жауап көзі емес, тексерілетін және сынға алынатын серіктес"],
    [I.grad, "Күтілетін нәтиже", "сыни ойлау деңгейінің артуы, ЖИ-ді жауапкершілікпен қолдану, тақырыпты терең түсіну"],
  ];
  concl.forEach((c, i) => {
    const y = 1.7 + i * 1.25;
    circleIcon(s, c[0], 0.7, y, 0.8, C.accent1);
    txt(s, [{ text: c[1] + ": ", options: { bold: true, color: C.accent2 } }, { text: c[2] }], { x: 1.75, y: y - 0.05, w: 7.6, h: 1.0, fontSize: 16, color: C.background1, valign: "middle" });
  });
  atomRings(s, 11.1, 3.6, 1.7, "FFFFFF", 40);
  txt(s, "Назарларыңызға рахмет! Сұрақтарыңызды күтемін", { x: 9.2, y: 6.2, w: 3.8, h: 0.7, fontSize: 16, bold: true, color: C.accent2, align: "center", valign: "middle" });
  s.addNotes("Қорытындылай келе: құзыреттілік тәсіл сыни ойлауды өзек етіп, оны пәндік құзырет пен ЖИ сауаттылығымен біріктіруге мүмкіндік береді. Слэйтер ережесі Атом құрылысы мен Периодтық заңды байланыстыратын қарапайым әрі тиімді құрал, ал оның шектеулері проблемалық жағдаяттарға таптырмас материал. ChatGPT пен Claude – дайын жауап беретін «оракул» емес, тексерілетін серіктес. Келесі кезеңде педагогикалық эксперимент жүргізіп, рубрика мен тест арқылы тиімділікті тексеремін. Назарларыңызға рахмет! Енді талқылауға көшейік. (≈2 мин + сұрақтар)");

  // ================= 20. REFERENCES =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қорытынды" });
  s.addText("Пайдаланылған әдебиеттер", { placeholder: "title" });
  const refs = [
    "Slater J.C. Atomic Shielding Constants // Physical Review. – 1930. – Vol. 36, № 1. – P. 57–64. – DOI: 10.1103/PhysRev.36.57",
    "Clementi E., Raimondi D.L. Atomic Screening Constants from SCF Functions // J. Chem. Phys. – 1963. – Vol. 38, № 11. – P. 2686–2689. – DOI: 10.1063/1.1733573",
    "Housecroft C.E., Sharpe A.G. Inorganic Chemistry. – 5th ed. – Harlow: Pearson, 2018.",
    "Ахметов Н.С. Общая и неорганическая химия. – М.: Высшая школа, 2001.",
    "Хуторской А.В. Ключевые компетенции как компонент личностно-ориентированной парадигмы образования // Народное образование. – 2003. – № 2. – С. 58–64.",
    "Зимняя И.А. Ключевые компетенции – новая парадигма результата образования // Высшее образование сегодня. – 2003. – № 5. – С. 34–42.",
    "Махмутов М.И. Проблемное обучение: основные вопросы теории. – М.: Педагогика, 1975.",
    "Лернер И.Я. Дидактические основы методов обучения. – М.: Педагогика, 1981.",
    "Facione P.A. Critical Thinking: A Statement of Expert Consensus for Purposes of Educational Assessment and Instruction (The Delphi Report). – Millbrae, CA: California Academic Press, 1990.",
    "Anderson L.W., Krathwohl D.R. (Eds.) A Taxonomy for Learning, Teaching, and Assessing. – New York: Longman, 2001.",
    "Clark T.M. Investigating the Use of an Artificial Intelligence Chatbot with General Chemistry Exam Questions // J. Chem. Educ. – 2023. – Vol. 100, № 5. – P. 1905–1916. – DOI: 10.1021/acs.jchemed.3c00027",
    "Emenike M.E., Emenike B.U. Was This Title Generated by ChatGPT? Considerations for Artificial Intelligence Text-Generation Software Programs for Chemists and Chemistry Educators // J. Chem. Educ. – 2023. – Vol. 100, № 4. – P. 1413–1418. – DOI: 10.1021/acs.jchemed.3c00063",
    "Kasneci E. et al. ChatGPT for good? On opportunities and challenges of large language models for education // Learning and Individual Differences. – 2023. – Vol. 103. – 102274. – DOI: 10.1016/j.lindif.2023.102274",
    "Zawacki-Richter O. et al. Systematic review of research on artificial intelligence applications in higher education – where are the educators? // Int. J. Educ. Technol. High. Educ. – 2019. – Vol. 16. – Art. 39. – DOI: 10.1186/s41239-019-0171-0",
    "UNESCO. Guidance for generative AI in education and research. – Paris: UNESCO, 2023.",
    "Қазақстан Республикасы Үкіметінің 2024 жылғы 24 шілдедегі № 592 қаулысы «Жасанды интеллектті дамытудың 2024–2029 жылдарға арналған тұжырымдамасын бекіту туралы».",
  ];
  const half = 8;
  [refs.slice(0, half), refs.slice(half)].forEach((col, ci) => {
    txt(s, col.map((r, i) => ({ text: r, options: { bullet: { type: "number", startAt: ci * half + 1 }, breakLine: i < col.length - 1 } })),
      { x: 0.6 + ci * 6.15, y: 1.4, w: 5.95, h: 5.5, fontSize: 10.5, paraSpaceAfter: 5 });
  });
  s.addNotes("Әдебиеттер тізімінде химиялық негіз (Слэйтер, Клементи–Раймонди, Housecroft, Ахметов), педагогикалық негіз (Хуторской, Зимняя, Махмутов, Лернер, Facione, Anderson–Krathwohl) және ЖИ бойынша заманауи зерттеулер (Clark, Emenike, Kasneci, Zawacki-Richter, ЮНЕСКО) бар. DOI сілтемелері арқылы барлығын тексеруге болады.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
