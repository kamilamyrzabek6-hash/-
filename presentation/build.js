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
  txt(s, "CHBBKTJI 6305 · СЕМИНАР САБАҒЫ · 7M01515 – ХИМИЯ", { x: 0.7, y: 0.8, w: 9, h: 0.4, fontSize: 13, bold: true, color: C.accent2, charSpacing: 1 });
  s.addText("Химиялық білім берудегі құзыреттілік тәсіл және жасанды интеллект", { placeholder: "title" });
  s.addText("«Атом құрылысы және Периодтық заң» тақырыбын Слэйтер ережесі арқылы оқыту мысалында", { placeholder: "body" });
  atomRings(s, 11.1, 3.3, 1.75, "FFFFFF", 40);
  txt(s, [
    { text: "Орындаған: ", options: { bold: true } }, { text: "Мырзабек Камила, 7M01515 – Химия, 2-курс магистранты", options: { breakLine: true } },
    { text: "Ғылыми жетекшісі: ", options: { bold: true } }, { text: "т.ғ.к., профессор Тапалова Анипа Сейдалықызы", options: { breakLine: true } },
    { text: "Қорқыт Ата атындағы Қызылорда университеті · 2026", options: {} },
  ], { x: 0.7, y: 5.7, w: 9.5, h: 1.2, fontSize: 15, color: C.background1, paraSpaceAfter: 4 });
  s.addNotes("Сәлеметсіздер ме! Бүгінгі семинар «Химиялық білім берудегі құзыреттілік тәсіл және жасанды интеллект» пәніне арналған. Пәннің негізгі идеяларын, яғни құзыреттілік тәсіл мен ЖИ-ді, бейорганикалық химияның нақты тақырыбы – Атом құрылысы және Периодтық заң – мысалында, Слэйтер ережесі арқылы көрсетемін. (≈1 мин)");

  // ================= 2. OUTLINE =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Семинар жоспары", { placeholder: "title" });
  const outline = [
    ["01", "Пәннің мақсаты мен міндеттері", "Семинар соларға қалай жауап береді"],
    ["02", "Химиялық білімдегі құзыреттілік тәсіл", "Тарихы, құрылымы, химиялық құзыреттер, мазмұнды саралау"],
    ["03", "Химиялық білімдегі ЖИ", "ЖИ негіздері, құралдар, промпттар, ЮНЕСКО құзыреттер жүйесі"],
    ["04", "Құзыреттілік тәсіл мен ЖИ-ді кіріктіру", "Кіріктіру моделі және сабақ құру алгоритмі"],
    ["05", "Қолданбалы мысал: Слэйтер ережесі", "Z*, периодтылық, проблемалық жағдаят, ЖИ жауабын тексеру"],
    ["06", "Мониторинг және қорытынды", "Бағалау рубрикасы, пән мақсатына қол жеткізу"],
  ];
  outline.forEach((o, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.15, y = 1.55 + row * 1.75;
    card(s, x, y, 5.95, 1.5, C.background2, "outline-card");
    txt(s, o[0], { x: x + 0.3, y: y + 0.3, w: 1.0, h: 0.9, fontSize: 36, bold: true, color: C.accent1, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, o[1], { x: x + 1.4, y: y + 0.28, w: 4.4, h: 0.45, fontSize: 18, bold: true, color: C.text2 });
    txt(s, o[2], { x: x + 1.4, y: y + 0.78, w: 4.4, h: 0.5, fontSize: 14, color: C.text1 });
  });
  s.addNotes("Семинар алты бөлімнен тұрады және пәннің силлабусындағы екі модульге сәйкес құрылған: 1-модуль – құзыреттілік түрлері мен сипаттамасы, 2-модуль – химиялық құзыреттіліктер және жасанды интеллект. Соңында осы екеуін бір сабақта біріктіретін нақты мысалды көрсетемін. (≈1 мин)");

  // ================= 3. COURSE AIM =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріспе" });
  s.addText("Пәннің мақсаты және семинардың құрылымы", { placeholder: "title" });
  card(s, 0.6, 1.45, 5.0, 4.0, C.text2, "aim");
  circleIcon(s, I.target, 0.9, 1.7, 0.8, C.accent2);
  txt(s, "Пәннің мақсаты (силлабус)", { x: 1.9, y: 1.8, w: 3.5, h: 0.6, fontSize: 16, bold: true, color: C.accent2, valign: "middle" });
  txt(s, "Химия пәні оқытушысының кәсіби қызметінің мақсаты – оқушыларды ойлануға, өз бетінше шешім қабылдауға, жаңалық ашуға, өз бетінше білім алуға, туындаған мәселелерді шешуге, қабылданған шешім үшін жауап беруге үйрету.", { x: 0.9, y: 2.7, w: 4.45, h: 2.6, fontSize: 15, italic: true, color: C.background1 });
  txt(s, "Пән міндеттері → семинар бөлімдері", { x: 6.0, y: 1.45, w: 6.7, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
  const tasks = [
    ["Құзыреттілік тұрғысынан қызмет құрылымын дамыту", "02 · 04"],
    ["Пәндік мазмұнды базалық және жоғары деңгейге саралау", "02 · 05"],
    ["Белсенді оқытудың құралдарын, әдістерін, технологияларын әзірлеу", "03 · 04 · 05"],
    ["Мониторингтің объективті жүйесін құру", "06"],
  ];
  tasks.forEach((t, i) => {
    const y = 1.95 + i * 0.9;
    card(s, 6.0, y, 6.7, 0.75, C.background2, "task");
    txt(s, String(i + 1), { x: 6.15, y, w: 0.4, h: 0.75, fontSize: 20, bold: true, color: C.accent1, valign: "middle", fontFace: THEME.headFontFace });
    txt(s, t[0], { x: 6.6, y, w: 4.6, h: 0.75, fontSize: 14, valign: "middle" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 11.3, y: y + 0.17, w: 1.25, h: 0.42, fill: { color: C.accent1 }, line: { color: C.accent1 }, rectRadius: 0.2, objectName: "chip" });
    txt(s, t[1], { x: 11.3, y: y + 0.17, w: 1.25, h: 0.42, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
  });
  card(s, 0.6, 5.7, 12.1, 1.1, C.background2, "on8");
  circleIcon(s, I.flag, 0.8, 5.85, 0.8, C.accent5);
  txt(s, [
    { text: "Оқыту нәтижесі (ON 8): ", options: { bold: true, color: C.text2 } },
    { text: "химиялық білім берудегі құзыреттілік тәсіл әдіснамасын және құзыреттілік тәсіл жағдайында химияны оқыту сапасын қамтамасыз ету тәсілдерін меңгерген, оларды кәсіби қызметте қолдана алады." },
  ], { x: 1.8, y: 5.75, w: 10.75, h: 1.0, fontSize: 14, valign: "middle" });
  s.addNotes("Семинардың өзегі – пәннің мақсаты. Силлабус бойынша химия оқытушысы оқушыларды ойлануға, өз бетінше шешім қабылдауға, мәселелерді шешуге және шешім үшін жауап беруге үйретуі тиіс. Пәннің төрт міндеті бар, оң жақта әр міндеттің семинардың қай бөлімінде ашылатыны көрсетілген. Мысалы, мазмұнды базалық және жоғары деңгейге саралау 2 және 5-бөлімдерде, ал мониторинг жүйесі 6-бөлімде қарастырылады. Төменде білім беру бағдарламасының ON 8 оқыту нәтижесі берілген: біз әдіснаманы меңгеріп қана қоймай, оны нақты сабақта қолдануымыз керек. (≈3 мин)");

  // ================= 4. HISTORY =================
  pres.addSection({ title: "Құзыреттілік тәсіл" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Құзыреттілік тәсіл" });
  s.addText("Құзыреттілік тәсілдің қалыптасуы", { placeholder: "title" });
  const tl = [
    ["1999", "Болон декларациясы", "Еуропалық жоғары білім кеңістігін құру"],
    ["2000", "TUNING жобасы", "Оқыту нәтижелерін құзыреттер арқылы сипаттау"],
    ["2005", "Берген", "ЕЖБК біліктілік шеңбері, Дублин дескрипторлары"],
    ["2010", "Қазақстан", "Болон процесіне 47-мүше болып қосылды"],
    ["2022", "МЖМБС", "Оқыту нәтижелеріне негізделген стандарт (ҒЖБМ № 2 бұйрық)"],
  ];
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 2.15, w: 11.5, h: 0, line: { color: C.accent6, width: 2 }, objectName: "timeline-line" });
  tl.forEach((t, i) => {
    const x = 0.6 + i * 2.47;
    s.addShape(pres.shapes.OVAL, { x: x + 0.85, y: 1.9, w: 0.5, h: 0.5, fill: { color: i === 3 ? C.accent2 : C.accent1 }, line: { color: "FFFFFF", width: 2 }, objectName: "tl-dot" });
    txt(s, t[0], { x, y: 1.3, w: 2.2, h: 0.5, fontSize: 22, bold: true, color: i === 3 ? C.accent2 : C.accent1, align: "center", fontFace: THEME.headFontFace });
    txt(s, t[1], { x, y: 2.55, w: 2.2, h: 0.4, fontSize: 15, bold: true, color: C.text2, align: "center" });
    txt(s, t[2], { x: x + 0.05, y: 2.95, w: 2.1, h: 0.9, fontSize: 13, align: "center" });
  });
  const defs = [
    ["Құзырет", "Білім алушының дайындығына қойылатын, алдын ала берілген әлеуметтік талап (норма)", C.background2, C.text2, C.text1],
    ["Құзыреттілік", "Тұлғаның сол талапты орындау қабілеті: білім + білік + тәжірибе + құндылықты жаңа жағдаятта қолдану", C.accent1, C.background1, C.background1],
  ];
  defs.forEach((d, i) => {
    const x = 0.6 + i * 6.15;
    card(s, x, 4.15, 5.95, 2.0, d[2], "def");
    circleIcon(s, i ? I.grad : I.book, x + 0.3, 4.45, 0.8, i ? C.text2 : C.accent1);
    txt(s, d[0], { x: x + 1.35, y: 4.4, w: 4.4, h: 0.45, fontSize: 18, bold: true, color: d[3] });
    txt(s, d[1], { x: x + 1.35, y: 4.9, w: 4.4, h: 1.15, fontSize: 14, color: d[4] });
  });
  txt(s, "Дереккөз: Хуторской А.В. (2003); Зимняя И.А. (2003); Gonzalez J., Wagenaar R. (2003); МЖМБС (2022)", { x: 0.6, y: 6.4, w: 12.1, h: 0.35, fontSize: 11, color: C.accent6 });
  s.addNotes("Құзыреттілік тәсілдің шығу тарихы аралық бақылау сұрақтарының бірі. 1999 жылғы Болон декларациясынан кейін 2000 жылы TUNING жобасы басталды: ол оқыту нәтижелерін құзыреттер арқылы сипаттауды ұсынды. 2005 жылы Бергенде Еуропалық жоғары білім кеңістігінің біліктілік шеңбері қабылданды, оның негізі – Дублин дескрипторлары. Қазақстан 2010 жылы Болон процесіне 47-мүше болып қосылды. Ал 2022 жылғы мемлекеттік стандарт оқыту нәтижелеріне негізделген. Төменде Хуторской бойынша екі ұғымның айырмашылығы: құзырет – талап, ал құзыреттілік – тұлғаның сол талапты орындау қабілеті. (≈3 мин)");

  // ================= 5. STRUCTURE =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Құзыреттілік тәсіл" });
  s.addText("Құзыреттілік құрылымы: химия мысалында", { placeholder: "title" });
  const comps = [
    [I.book, "Когнитивтік (білім)", "Атом құрылысы, электрондық конфигурация, Слэйтер ережесі, Z* ұғымы", C.accent1],
    [I.puzzle, "Әрекеттік (білік)", "Z*-ті есептеу, иондану энергиясының графигін талдау, болжам жасау", C.accent5],
    [I.bulb, "Мотивациялық-құндылықтық", "Танымдық қызығушылық, «Неге?» деген сұрақ, ЖИ-ді адал қолдану", C.accent4],
    [I.sync, "Рефлексивтік", "Өз шешімін және ЖИ жауабын бағалау, модельдің шектеулерін мойындау", C.accent3],
  ];
  comps.forEach((c, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = col ? 7.9 : 0.6, y = 1.5 + row * 2.7;
    card(s, x, y, 4.8, 2.45, C.background2, "comp-card");
    circleIcon(s, c[0], x + 0.25, y + 0.25, 0.75, c[3]);
    txt(s, c[1], { x: x + 1.15, y: y + 0.25, w: 3.5, h: 0.75, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
    txt(s, c[2], { x: x + 0.25, y: y + 1.15, w: 4.35, h: 1.2, fontSize: 14 });
  });
  s.addShape(pres.shapes.OVAL, { x: 5.6, y: 2.75, w: 2.15, h: 2.15, fill: { color: C.text2 }, line: { color: C.accent2, width: 3 }, objectName: "center" });
  txt(s, "Құзырет-тілік", { x: 5.6, y: 2.75, w: 2.15, h: 2.15, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle" });
  s.addNotes("Құзыреттілік құрылымы да аралық бақылау сұрағы. Оны төрт компонент арқылы қарастырамын және әрқайсысын біздің тақырыпқа байланыстырамын. Когнитивтік компонент – білім: Слэйтер ережесі, Z*. Әрекеттік компонент – білік: Z*-ті есептеу, графикті талдау. Мотивациялық-құндылықтық компонент – «неге олай?» деген қызығушылық және ЖИ-ді адал қолдану. Рефлексивтік компонент – өз шешімін және ЖИ жауабын бағалау. Төртеуі бірге болғанда ғана құзыреттілік қалыптасады. (≈2,5 мин)");

  // ================= 6. CHEMICAL COMPETENCIES =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Құзыреттілік тәсіл" });
  s.addText("Химияны оқытудағы негізгі құзыреттер", { placeholder: "title" });
  const hd = (t, a) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 }, align: a || "left" } });
  const rows = [
    [hd("Құзырет"), hd("«Атом құрылысы» тақырыбында"), hd("Пән мақсатымен байланысы")],
    [{ text: "Пәндік (химиялық)", options: { bold: true } }, "Z* есептеу, периодтық заңдылықтарды түсіндіру", "мәселелерді шешу"],
    [{ text: "Сыни ойлау", options: { bold: true, color: H.accent3 } }, "Деректерді талдау, ЖИ жауабын тексеру, модель шектеулерін бағалау", "ойлану, шешім үшін жауап беру"],
    [{ text: "Ақпараттық-цифрлық (ЖИ сауаттылығы)", options: { bold: true } }, "Промпт құрастыру, ЖИ жауабын дереккөзбен салыстыру", "өз бетінше білім алу"],
    [{ text: "Зерттеушілік", options: { bold: true } }, "Болжам ұсыну, есептеу арқылы тексеру", "жаңалық ашу"],
    [{ text: "Коммуникативтік", options: { bold: true } }, "Топта талқылау, шешімді дәлелдеп қорғау", "өз бетінше шешім қабылдау"],
  ];
  s.addTable(rows, { x: 0.6, y: 1.5, w: 8.1, colW: [2.4, 3.4, 2.3], fontSize: 13, color: H.dk1, border: { type: "solid", pt: 0.75, color: "C9DADF" }, fill: { color: "FFFFFF" }, valign: "middle", rowH: 0.75, margin: 0.08, objectName: "competency-table" });
  const tri = [[9.55, 1.6, C.accent1], [11.1, 1.6, C.accent5], [10.32, 2.9, C.accent3]];
  tri.forEach(t => s.addShape(pres.shapes.OVAL, { x: t[0] - 0.75, y: t[1], w: 2.1, h: 2.1, fill: { color: t[2], transparency: 25 }, line: { color: "FFFFFF", width: 1 }, objectName: "triad" }));
  txt(s, "Пәндік", { x: 8.85, y: 2.2, w: 1.6, h: 0.4, fontSize: 14, bold: true, color: C.background1, align: "center" });
  txt(s, "ЖИ сауат-тылығы", { x: 10.95, y: 2.1, w: 1.3, h: 0.6, fontSize: 13, bold: true, color: C.background1, align: "center" });
  txt(s, "Сыни ойлау", { x: 9.65, y: 4.25, w: 1.8, h: 0.4, fontSize: 14, bold: true, color: C.background1, align: "center" });
  card(s, 8.95, 5.25, 3.75, 1.55, C.background2, "rec");
  txt(s, [{ text: "Ұсыныс: ", options: { bold: true, color: C.accent3 } }, { text: "үш құзыретті бір тапсырмада қатар дамыту. Сыни ойлау – пән мақсатының өзегі." }], { x: 9.15, y: 5.35, w: 3.4, h: 1.4, fontSize: 13, valign: "middle" });
  s.addNotes("«Химияны оқытудағы негізгі құзыреттіліктер» деген сұрақ та бақылау сұрақтарының бірі. Кестеде бес құзырет берілген, әрқайсысы біздің тақырыппен және пәннің мақсатымен байланыстырылған. Мысалы, сыни ойлау пән мақсатындағы «ойлануға» және «шешім үшін жауап беруге» сәйкес келеді. Менің ұсынысым: пәндік құзыретті, сыни ойлауды және ЖИ сауаттылығын бір тапсырмада қатар дамыту. Студент Z*-ті өзі есептейді, ЖИ-ден сұрайды, сосын екеуін салыстырып, қатені табады. Сыни ойлау құрылымын P. Facione (1990) бойынша аламын: интерпретациялау, талдау, бағалау, қорытынды шығару, түсіндіру, өзін-өзі реттеу. (≈3 мин)");

  // ================= 7. DIFFERENTIATION =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Құзыреттілік тәсіл" });
  s.addText("Пәндік мазмұнды деңгейлерге саралау", { placeholder: "title" });
  txt(s, "Пән міндеті 2 · Блумның қайта қаралған таксономиясы бойынша (Anderson, Krathwohl, 2001)", { x: 0.6, y: 1.25, w: 12.1, h: 0.4, fontSize: 14, color: C.accent6 });
  const lv = [
    ["Базалық деңгей", "Есте сақтау · түсіну", ["Элементтің электрондық конфигурациясын жазу", "Период пен топ бойынша радиус пен иондану энергиясының өзгеруін атау"], C.accent5, 0],
    ["Орта деңгей", "Қолдану · талдау", ["Слэйтер ережесімен Z*-ті есептеу", "Z* арқылы иондану энергиясының графигін түсіндіру"], C.accent1, 1],
    ["Жоғары деңгей", "Бағалау · жасау", ["Fe²⁺ мәселесін шешу, B мен O аномалиясын талдау", "ЖИ жауабын сынау, өз проблемалық есебін құрастыру"], C.accent3, 2],
  ];
  lv.forEach((l) => {
    const i = l[4], x = 0.6 + i * 4.1, h = 3.0 + i * 0.7, y = 6.75 - h;
    card(s, x, y, 3.9, h, C.background2, "level");
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.9, h: 0.95, fill: { color: l[3] }, line: { color: l[3] }, objectName: "level-head" });
    txt(s, l[0], { x: x + 0.25, y: y + 0.08, w: 3.4, h: 0.45, fontSize: 18, bold: true, color: C.background1 });
    txt(s, l[1], { x: x + 0.25, y: y + 0.5, w: 3.4, h: 0.35, fontSize: 13, color: C.background1, italic: true });
    txt(s, l[2].map((b, k) => ({ text: b, options: { bullet: true, breakLine: k < l[2].length - 1 } })), { x: x + 0.25, y: y + 1.15, w: 3.45, h: h - 1.3, fontSize: 14, paraSpaceAfter: 8 });
  });
  s.addShape(pres.shapes.RIGHT_ARROW, { x: 0.6, y: 1.8, w: 2.2, h: 0.35, fill: { color: C.accent2 }, line: { color: C.accent2 }, objectName: "arrow" });
  txt(s, "күрделілік артады", { x: 2.95, y: 1.75, w: 3, h: 0.45, fontSize: 13, italic: true, color: C.accent6, valign: "middle" });
  s.addNotes("Пәннің екінші міндеті – мазмұнды базалық және жоғары деңгейге саралау. Мұны Блумның қайта қаралған таксономиясы бойынша үш деңгейге бөлдім. Базалық деңгейде студент конфигурацияны жазып, заңдылықтарды атайды. Орта деңгейде Слэйтер ережесімен Z*-ті есептеп, графикті түсіндіреді. Жоғары деңгейде проблемалық жағдаятты шешеді, ЖИ жауабын сынайды және өзі есеп құрастырады. Бұл силлабустағы «Әр түрлі деңгейлі тапсырмалар» тақырыбына да сәйкес. Бір тақырыпта әр студент өз деңгейінен бастап, жоғарыға көтеріле алады. (≈3 мин)");

  // ================= 8. AI BASICS =================
  pres.addSection({ title: "Химиялық білімдегі ЖИ" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Химиялық білімдегі ЖИ" });
  s.addText("Жасанды интеллект негіздері", { placeholder: "title" });
  const nest = [["Жасанды интеллект (ЖИ)", 3.0, C.accent5], ["Машиналық оқыту (ML)", 2.35, C.accent1], ["Тереңдетілген оқыту (DL)", 1.7, C.accent4], ["Тілдік модельдер (NLP)", 1.05, C.accent2]];
  const cx = 3.6, cy = 4.2;
  nest.forEach((n, i) => {
    s.addShape(pres.shapes.OVAL, { x: cx - n[1] * 1.0, y: cy - n[1] * 0.85, w: n[1] * 2.0, h: n[1] * 1.7, fill: { color: n[2], transparency: i === 3 ? 0 : 15 }, line: { color: "FFFFFF", width: 1.5 }, objectName: "nest" });
  });
  [["Жасанды интеллект (ЖИ)", 1.8], ["Машиналық оқыту (ML)", 2.33], ["Тереңдетілген оқыту (DL)", 2.88]].forEach((t, i) => txt(s, t[0], { x: cx - 2, y: t[1], w: 4, h: 0.35, fontSize: 13, bold: true, color: C.background1, align: "center" }));
  txt(s, [{ text: "Тілдік модельдер", options: { bold: true, breakLine: true } }, { text: "ChatGPT · Claude", options: { fontSize: 13 } }], { x: cx - 1.0, y: cy - 0.45, w: 2.0, h: 0.9, fontSize: 14, color: C.text1, align: "center", valign: "middle" });
  txt(s, "ЖИ түрлері", { x: 7.4, y: 1.45, w: 5.3, h: 0.4, fontSize: 18, bold: true, color: C.text2 });
  const types = [
    ["Тар (narrow) ЖИ", "Бір міндетке арналған. Қазіргі жүйелердің бәрі, ChatGPT пен Claude да, осы түрге жатады", C.accent1],
    ["Жалпы (general) ЖИ", "Адам сияқты кез келген зияткерлік міндетті орындай алатын жүйе. Әзірге гипотезалық", C.accent5],
    ["Супер интеллект", "Адам интеллектінен асып түсетін теориялық деңгей", C.accent6],
  ];
  types.forEach((t, i) => {
    const y = 1.95 + i * 1.35;
    card(s, 7.4, y, 5.3, 1.2, C.background2, "type");
    s.addShape(pres.shapes.OVAL, { x: 7.6, y: y + 0.38, w: 0.45, h: 0.45, fill: { color: t[2] }, line: { color: t[2] }, objectName: "dot" });
    txt(s, t[0], { x: 8.25, y: y + 0.1, w: 4.3, h: 0.35, fontSize: 15, bold: true, color: C.text2 });
    txt(s, t[1], { x: 8.25, y: y + 0.45, w: 4.3, h: 0.7, fontSize: 13 });
  });
  card(s, 7.4, 6.0, 5.3, 0.8, C.text2, "note");
  txt(s, "Тілдік модель мәтіннің ықтимал жалғасын болжайды. Ол химиялық есептеулерді тексермейді, сондықтан қателесуі мүмкін.", { x: 7.6, y: 6.0, w: 5.0, h: 0.8, fontSize: 13, color: C.background1, valign: "middle" });
  s.addNotes("Силлабустағы ЖИ негіздері тақырыбына қысқаша тоқталамын. Сол жақта ұғымдардың бір-біріне енуі көрсетілген: жасанды интеллект ішінде машиналық оқыту, оның ішінде тереңдетілген оқыту, ал ChatGPT пен Claude – табиғи тілді өңдейтін үлкен тілдік модельдер. ЖИ-дің үш түрі бар: тар, жалпы және супер интеллект. Қазіргі жүйелердің бәрі тар ЖИ-ге жатады. Ең маңызды әдістемелік қорытынды: тілдік модель химиялық есептеуді тексермейді, ол ықтимал мәтін құрастырады. Сондықтан оның жауабын тексеру керек, ал бұл сыни ойлауды дамытуға тамаша мүмкіндік береді. (≈3 мин)");

  // ================= 9. TOOLS =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Химиялық білімдегі ЖИ" });
  s.addText("Химиялық білімдегі цифрлық құралдар", { placeholder: "title" });
  const tools = [
    [I.robot, "ChatGPT (OpenAI)", "Генеративтік ЖИ", "Сократтық диалог, есептеуді тексеру, проблемалық жағдаят ұсыну", C.accent1],
    [I.robot, "Claude (Anthropic)", "Генеративтік ЖИ", "Түсіндірмелерді салыстыру, модельдің шектеулерін талқылау", C.accent1],
    [I.laptop, "DeepChem", "ML кітапханасы", "Химиядағы машиналық оқыту: зерттеушілік деңгейге арналған", C.accent5],
    [I.atom, "PhET «Build an Atom»", "Симуляция", "Протон, нейтрон, электрон арқылы атомды «құрастыру»", C.accent4],
    [I.flask, "MolView", "Визуализация", "Молекулалар мен орбитальдардың 3D моделі", C.accent4],
    [I.eye, "ChemCollective", "Виртуалды зертхана", "Виртуалды тәжірибелер және контексті тапсырмалар", C.accent4],
  ];
  tools.forEach((t, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.1, y = 1.45 + row * 2.4;
    card(s, x, y, 3.9, 2.2, C.background2, "tool");
    circleIcon(s, t[0], x + 0.25, y + 0.25, 0.7, t[4]);
    txt(s, t[1], { x: x + 1.1, y: y + 0.22, w: 2.7, h: 0.4, fontSize: 15, bold: true, color: C.text2 });
    txt(s, t[2], { x: x + 1.1, y: y + 0.6, w: 2.7, h: 0.35, fontSize: 12, italic: true, color: t[4] === C.accent4 ? C.accent4 : C.accent1 });
    txt(s, t[3], { x: x + 0.25, y: y + 1.1, w: 3.45, h: 1.0, fontSize: 14 });
  });
  txt(s, [{ text: "Сыни ескерту: ", options: { bold: true, color: C.accent3 } }, { text: "PhET, MolView, ChemCollective – ЖИ емес, цифрлық модельдеу құралдары. Олар ЖИ-мен бірге қолданылады. Құралдың түрін ажырата білудің өзі – цифрлық сауаттылық." }], { x: 0.6, y: 6.3, w: 12.1, h: 0.6, fontSize: 13, valign: "middle" });
  s.addNotes("Силлабуста ұсынылған құралдарды түрлері бойынша топтастырдым. ChatGPT пен Claude – генеративтік ЖИ, менің жұмысымның негізгі құралдары. DeepChem – химиядағы машиналық оқыту кітапханасы, ол зерттеушілік деңгейге арналған. PhET, MolView және ChemCollective – симуляция, визуализация және виртуалды зертхана. Атом құрылысы тақырыбына PhET-тің «Build an Atom» симуляциясы өте ыңғайлы. Бір маңызды ескерту: бұл үш құрал ЖИ емес. Оларды жиі шатастырады, сондықтан құралдың түрін ажырата білудің өзі цифрлық сауаттылықтың бөлігі. (≈2,5 мин)");

  // ================= 10. ROLES + PROMPTS =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Химиялық білімдегі ЖИ" });
  s.addText("ЖИ-мен жұмыс: рөлдер мен промпттар", { placeholder: "title" });
  const pr = [
    [I.comments, "Сократтық тәлімгер", "«Сен химия оқытушысысың. Дайын жауап берме. Cl атомының 3p электроны үшін Z*-ті есептеуге жетелейтін сұрақтарды бір-бірлеп қой.»"],
    [I.check, "Тексеруші", "«Менің есептеуім: S = 2×1,00 + 8×0,85 + 6×0,35. Қай қадам қате? Тек себебін түсіндір, түзетуді өзім жасаймын.»"],
    [I.warn, "Сынға алынатын серіктес", "«Слэйтер ережесі B мен O-ның иондану энергиясындағы ауытқуды түсіндіре ала ма? Модельдің шектеулерін ата.»"],
  ];
  pr.forEach((p, i) => {
    const y = 1.45 + i * 1.8;
    card(s, 0.6, y, 7.7, 1.6, C.background2, "prompt");
    circleIcon(s, p[0], 0.8, y + 0.2, 0.65, C.accent1);
    txt(s, p[1], { x: 1.65, y: y + 0.15, w: 6.5, h: 0.4, fontSize: 15, bold: true, color: C.text2 });
    txt(s, p[2], { x: 1.65, y: y + 0.55, w: 6.5, h: 1.0, fontSize: 13, italic: true });
  });
  card(s, 8.6, 1.45, 4.1, 5.3, C.text2, "risk");
  txt(s, "Тәуекелдер және шешімдер", { x: 8.85, y: 1.65, w: 3.6, h: 0.4, fontSize: 17, bold: true, color: C.accent2 });
  txt(s, [
    { text: "«Галлюцинация»: қате сандар мен ойдан шығарылған дереккөздер", options: { bullet: true, breakLine: true } },
    { text: "→ анықтамалық деректермен салыстыру", options: { color: H.accent2, breakLine: true } },
    { text: "Ойлауды ЖИ-ге тапсыру, тәуелділік", options: { bullet: true, breakLine: true } },
    { text: "→ алдымен өз болжамы, содан кейін ЖИ", options: { color: H.accent2, breakLine: true } },
    { text: "Академиялық адалдық", options: { bullet: true, breakLine: true } },
    { text: "→ ЖИ-мен диалогты есепке қоса тапсыру", options: { color: H.accent2 } },
  ], { x: 8.85, y: 2.2, w: 3.65, h: 3.9, fontSize: 14, color: C.background1, paraSpaceAfter: 6 });
  txt(s, "UNESCO (2023); Kasneci et al. (2023); Clark (2023)", { x: 8.85, y: 6.25, w: 3.65, h: 0.35, fontSize: 11, color: C.background2 });
  s.addNotes("ЖИ-ге үш дидактикалық рөл беремін және әрқайсысына промпт үлгісін ұсынамын. Біріншісі – Сократ әдісімен сұрақ қоятын тәлімгер: ол дайын жауап бермейді. Екіншісі – тексеруші: ол қатенің қай жерде екенін ғана көрсетеді, ал түзетуді студент өзі жасайды. Үшіншісі – сынға алынатын серіктес: студент ЖИ-мен модельдің шектеулері туралы пікірталас жүргізеді. Оң жақта тәуекелдер және оларды шешу жолдары. Clark (2023) ChatGPT-дің химия сұрақтарына сенімді, бірақ қате жауап беретінін көрсетіп, оның жауаптарын студенттерге сынату тапсырмаларын ұсынды. (≈3 мин)");

  // ================= 11. UNESCO =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Химиялық білімдегі ЖИ" });
  s.addText("ЮНЕСКО: ЖИ құзыреттер жүйесі", { placeholder: "title" });
  txt(s, "Мұғалімдерге арналған ЖИ құзыреттер жүйесі (UNESCO, 2024): 5 бағыт, 15 құзырет, 3 даму деңгейі", { x: 0.6, y: 1.25, w: 12.1, h: 0.4, fontSize: 14, color: C.accent6 });
  const un = [
    [I.users, "Адамға бағдарланған ойлау", "Шешімді ЖИ емес, адам қабылдайды", C.accent1],
    [I.scale, "ЖИ этикасы", "Адалдық, деректер қауіпсіздігі, авторлық", C.accent3],
    [I.laptop, "ЖИ негіздері мен қолданылуы", "Тілдік модель қалай жұмыс істейді, шегі қандай", C.accent5],
    [I.grad, "ЖИ педагогикасы", "ЖИ-ді проблемалық сабаққа кіріктіру", C.accent4],
    [I.sync, "Кәсіби даму үшін ЖИ", "Тапсырма әзірлеу, өз тәжірибесін талдау", C.accent2],
  ];
  un.forEach((u, i) => {
    const x = 0.6 + i * 2.47;
    card(s, x, 1.85, 2.27, 3.2, C.background2, "un");
    circleIcon(s, u[0], x + 0.73, 2.05, 0.8, u[3]);
    txt(s, u[1], { x: x + 0.15, y: 3.0, w: 1.97, h: 0.85, fontSize: 14, bold: true, color: C.text2, align: "center", valign: "middle" });
    txt(s, u[2], { x: x + 0.15, y: 3.9, w: 1.97, h: 1.05, fontSize: 13, align: "center" });
  });
  const lvl = ["Меңгеру", "Тереңдету", "Жасау"];
  lvl.forEach((l, i) => {
    const x = 0.6 + i * 4.1;
    s.addShape(pres.shapes.CHEVRON, { x, y: 5.35, w: 3.9, h: 0.75, fill: { color: [C.accent6, C.accent1, C.text2][i] }, line: { color: "FFFFFF" }, objectName: "level-chevron" });
    txt(s, `${i + 1}-деңгей · ${l}`, { x: x + 0.4, y: 5.35, w: 3.1, h: 0.75, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle" });
  });
  txt(s, "Болашақ химия мұғалімі үшін ЖИ сауаттылығы – пәндік құзыретпен қатар тұратын кәсіби құзырет.", { x: 0.6, y: 6.3, w: 12.1, h: 0.5, fontSize: 14, italic: true, color: C.text2 });
  s.addNotes("Магистранттар – болашақ оқытушылар, сондықтан ЮНЕСКО-ның 2024 жылғы мұғалімдерге арналған ЖИ құзыреттер жүйесін ұсынамын. Онда бес бағыт бар: адамға бағдарланған ойлау, ЖИ этикасы, ЖИ негіздері мен қолданылуы, ЖИ педагогикасы және кәсіби даму үшін ЖИ. Барлығы 15 құзырет, үш даму деңгейінде: меңгеру, тереңдету және жасау. Біздің сабақ әсіресе «ЖИ педагогикасы» мен «адамға бағдарланған ойлау» бағыттарын дамытады: шешімді ЖИ емес, студент қабылдайды. Бұл пән мақсатындағы «шешім үшін жауап беру» талабына тікелей сәйкес келеді. (≈2,5 мин)");

  // ================= 12. INTEGRATION MODEL =================
  pres.addSection({ title: "Кіріктіру" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріктіру" });
  s.addText("Құзыреттілік тәсіл мен ЖИ-ді кіріктіру моделі", { placeholder: "title" });
  const blocks = [
    ["МАҚСАТ", "Пән мақсаты: ойлану, шешім қабылдау, мәселе шешу", C.accent3, I.target],
    ["МАЗМҰН", "Атом құрылысы · Периодтық заң · Слэйтер ережесі (3 деңгей)", C.accent1, I.atom],
    ["ӘДІС", "Белсенді оқыту: проблемалық оқыту", C.accent5, I.puzzle],
    ["ҚҰРАЛ", "ChatGPT, Claude + PhET симуляциясы", C.accent4, I.robot],
    ["МОНИТОРИНГ", "Рубрика, тест, ЖИ-диалог портфолиосы", C.accent2, I.flag],
  ];
  blocks.forEach((b, i) => {
    const x = 0.6 + i * 2.475;
    card(s, x, 1.55, 2.2, 3.15, b[2], "model-block");
    circleIcon(s, b[3], x + 0.7, 1.75, 0.8, C.text2);
    txt(s, b[0], { x: x + 0.12, y: 2.7, w: 1.96, h: 0.4, fontSize: 15, bold: true, color: i === 4 ? C.text1 : C.background1, align: "center" });
    txt(s, b[1], { x: x + 0.12, y: 3.15, w: 1.96, h: 1.45, fontSize: 13, color: i === 4 ? C.text1 : C.background1, align: "center" });
    if (i < 4) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 2.24, y: 2.95, w: 0.2, h: 0.35, fill: { color: C.accent6 }, line: { color: C.accent6 }, objectName: "arrow" });
  });
  card(s, 0.6, 4.95, 12.1, 1.85, C.background2, "why-topic");
  circleIcon(s, I.atom, 0.85, 5.3, 1.0, C.accent1);
  txt(s, [
    { text: "Неге «Атом құрылысы және Периодтық заң»? ", options: { bold: true, color: C.text2, breakLine: true } },
    { text: "Бұл бейорганикалық химияның іргетасы. Слэйтер ережесі электрондық конфигурацияны периодтық қасиеттермен қарапайым есептеу арқылы байланыстырады. Ал оның шектеулері проблемалық жағдаят туғызуға таптырмас материал (силлабус: МӨЖ 11 «Бейорганикалық химияны оқытудағы ЖИ маңызы»)." },
  ], { x: 2.1, y: 5.05, w: 10.4, h: 1.65, fontSize: 15, valign: "middle" });
  s.addNotes("Бұл слайд пәннің екі модулін біріктіреді. Модель бес компоненттен тұрады және әрқайсысы пән міндеттерінің біріне сәйкес келеді. Мақсат – пән мақсатының өзі. Мазмұн – үш деңгейге сараланған Атом құрылысы тақырыбы (2-міндет). Әдіс – проблемалық оқыту, белсенді оқытудың бір түрі (3-міндет). Құрал – ChatGPT, Claude және PhET. Мониторинг – рубрика, тест және портфолио (4-міндет). Тақырыпты таңдау себебі: Слэйтер ережесі қарапайым есептеу арқылы заңдылықты түсіндіреді, ал шектеулері проблемалық жағдаятқа арқау болады. Бұл силлабустағы МӨЖ 11 тақырыбына да тікелей сәйкес. (≈3 мин)");

  // ================= 13. LESSON ALGORITHM =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Кіріктіру" });
  s.addText("Сабақ алгоритмі (50 минут)", { placeholder: "title" });
  const st = [
    ["5", "Проблемалық жағдаят", "«Неге олай?» деген қайшылық туғызу", false],
    ["7", "Мәселе және болжам", "Топта нұсқалар ұсыну, ЖИ-сіз", false],
    ["13", "Өз бетінше шешу", "Слэйтер бойынша Z* есептеу, PhET", false],
    ["12", "ЖИ-мен тексеру", "ChatGPT / Claude жауабын салыстыру", true],
    ["8", "Дәлелдеу", "Топ шешімін қорғау", false],
    ["5", "Рефлексия", "Модельдің шегі, өзін бағалау", false],
  ];
  st.forEach((t, i) => {
    const x = 0.6 + i * 2.05;
    const col = t[3] ? C.accent2 : C.accent1;
    s.addShape(pres.shapes.OVAL, { x: x + 0.5, y: 1.55, w: 0.95, h: 0.95, fill: { color: col }, line: { color: "FFFFFF" }, objectName: "step" });
    txt(s, [{ text: t[0], options: { fontSize: 22, bold: true, breakLine: true } }, { text: "мин", options: { fontSize: 11 } }], { x: x + 0.5, y: 1.55, w: 0.95, h: 0.95, color: t[3] ? C.text1 : C.background1, align: "center", valign: "middle" });
    if (i < 5) s.addShape(pres.shapes.LINE, { x: x + 1.55, y: 2.02, w: 1.85, h: 0, line: { color: C.accent6, width: 1.5, dashType: "dash" }, objectName: "connector" });
    txt(s, t[1], { x, y: 2.65, w: 1.95, h: 0.65, fontSize: 14, bold: true, color: C.text2, align: "center" });
    txt(s, t[2], { x: x + 0.05, y: 3.3, w: 1.85, h: 0.9, fontSize: 13, align: "center" });
  });
  const principles = [
    [I.key, "Алдымен өзі, содан кейін ЖИ", "Алғашқы 25 минутта ЖИ қолданылмайды"],
    [I.users, "Топтық жұмыс", "3–4 адамнан, рөлдері бөлінеді"],
    [I.check, "Дәлел міндетті", "Әр тұжырым есептеумен немесе дереккөзбен расталады"],
  ];
  principles.forEach((p, i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 4.55, 3.9, 1.65, C.background2, "principle");
    circleIcon(s, p[0], x + 0.25, 4.8, 0.7, C.accent5);
    txt(s, p[1], { x: x + 1.1, y: 4.7, w: 2.7, h: 0.6, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    txt(s, p[2], { x: x + 1.1, y: 5.3, w: 2.7, h: 0.8, fontSize: 13 });
  });
  txt(s, "Негізі: Махмутов М.И. (1975), Лернер И.Я. (1981): проблемалық оқыту кезеңдері · силлабус: МОӨЖ 3 «Құзыреттілік тәсілдегі сабақты құру алгоритмі»", { x: 0.6, y: 6.45, w: 12.1, h: 0.35, fontSize: 11, color: C.accent6 });
  s.addNotes("Силлабустағы МОӨЖ 3 «Құзыреттілік тәсілдегі сабақты құру алгоритмі» тақырыбына сәйкес, Махмутов пен Лернердің проблемалық оқыту кезеңдері негізінде 50 минуттық сабақ құрдым. Алты кезең бар. ЖИ тек 4-кезеңде, тексеру үшін қосылады. Үш негізгі принцип: алдымен студент өзі ойлайды, сосын ЖИ қолданылады; жұмыс топпен жүреді; әр тұжырым есептеумен немесе дереккөзбен дәлелденеді. Бұл пән мақсатындағы «өз бетінше шешім қабылдау» талабын тікелей жүзеге асырады. (≈3 мин)");

  // ================= 14. SLATER RULES + EXAMPLE =================
  pres.addSection({ title: "Қолданбалы мысал" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қолданбалы мысал" });
  s.addText("Слэйтер ережесі: ядроның тиімді заряды", { placeholder: "title" });
  card(s, 0.6, 1.45, 6.0, 1.35, C.text2, "formula");
  txt(s, "Z* = Z − S", { x: 0.8, y: 1.5, w: 2.9, h: 1.25, fontSize: 36, bold: true, color: C.background1, valign: "middle", fontFace: THEME.headFontFace });
  txt(s, [
    { text: "Z* – ядроның тиімді заряды", options: { breakLine: true } },
    { text: "Z – ядро заряды", options: { breakLine: true } },
    { text: "S – экрандау (бүркеу) тұрақтысы", options: {} },
  ], { x: 3.7, y: 1.55, w: 2.8, h: 1.15, fontSize: 13, color: C.background2, valign: "middle" });
  const sr = [
    [hd("Экрандаушы электрондар"), hd("ns, np үшін", "center"), hd("nd, nf үшін", "center")],
    ["Оң жақтағы топтар", { text: "0", options: { align: "center" } }, { text: "0", options: { align: "center" } }],
    ["Сол топтағы басқа электрондар", { text: "0,35 (1s: 0,30)", options: { align: "center" } }, { text: "0,35", options: { align: "center" } }],
    ["(n − 1) қабат", { text: "0,85", options: { align: "center", bold: true, color: H.accent1 } }, { text: "1,00", options: { align: "center" } }],
    ["(n − 2) және одан төмен", { text: "1,00", options: { align: "center" } }, { text: "1,00", options: { align: "center" } }],
  ];
  s.addTable(sr, { x: 0.6, y: 3.0, w: 6.0, colW: [2.7, 1.8, 1.5], fontSize: 13, color: H.dk1, border: { type: "solid", pt: 0.75, color: "C9DADF" }, fill: { color: "FFFFFF" }, valign: "middle", rowH: 0.55, margin: 0.07, objectName: "slater-table" });
  txt(s, "Топтар: (1s)(2s,2p)(3s,3p)(3d)(4s,4p)(4d)(4f)(5s,5p)… · Slater, 1930", { x: 0.6, y: 5.9, w: 6.0, h: 0.6, fontSize: 12, color: C.accent6 });
  // example
  txt(s, "Мысал: Cl (Z = 17), 3p электроны", { x: 7.0, y: 1.45, w: 5.7, h: 0.45, fontSize: 17, bold: true, color: C.text2 });
  const steps = [
    ["(1s²)", "2 × 1,00 = 2,00", C.accent5, false],
    ["(2s² 2p⁶)", "8 × 0,85 = 6,80", C.accent1, false],
    ["(3s² 3p⁵)", "6 × 0,35 = 2,10", C.accent2, true],
  ];
  steps.forEach((t, i) => {
    const y = 2.05 + i * 0.85;
    card(s, 7.0, y, 1.9, 0.68, t[2], "grp");
    txt(s, t[0], { x: 7.0, y, w: 1.9, h: 0.68, fontSize: 15, bold: true, color: t[3] ? C.text1 : C.background1, align: "center", valign: "middle" });
    txt(s, t[1], { x: 9.1, y, w: 3.5, h: 0.68, fontSize: 18, valign: "middle" });
  });
  card(s, 7.0, 4.7, 5.7, 1.35, C.text2, "result");
  txt(s, [
    { text: "S = 2,00 + 6,80 + 2,10 = 10,90", options: { breakLine: true } },
    { text: "Z* = 17 − 10,90 = 6,10", options: { bold: true, color: H.accent2, fontSize: 24 } },
  ], { x: 7.25, y: 4.75, w: 5.3, h: 1.25, fontSize: 17, color: C.background1, valign: "middle" });
  txt(s, "Ішкі электрондар ядро зарядын «бүркейді»: сыртқы электрон тек Z* бөлігін «сезеді».", { x: 7.0, y: 6.15, w: 5.7, h: 0.6, fontSize: 13, italic: true, color: C.text2 });
  s.addNotes("Енді қолданбалы мысалға көшеміз. Слэйтер ережесін 1930 жылы Джон Слэйтер ұсынған: Z* = Z − S. Кестеде экрандау коэффициенттері берілген: оң жақтағы электрондар 0, сол топтағы электрондар 0,35, (n−1) қабат 0,85, одан ішкі қабаттар 1,00. Хлордың 3p электроны үшін: 1s қабатының 2 электроны 2×1,00, екінші қабаттың 8 электроны 8×0,85, үшінші қабаттың қалған 6 электроны 6×0,35. S = 10,90, Z* = 6,10. Клементи мен Раймондидің (1963) өздігінен келісілген өріс әдісімен есептеген мәні 6,12, яғни сәйкестік өте жақсы. Бұл – орта деңгейдегі тапсырма. (≈3 мин)");

  // ================= 13. PERIODICITY CHARTS =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қолданбалы мысал" });
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
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қолданбалы мысал" });
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

  // ================= 16. CRITIQUE AI =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Қолданбалы мысал" });
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

  pres.addSection({ title: "Мониторинг және қорытынды" });
  // ================= 18. RUBRIC =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Мониторинг және қорытынды" });
  s.addText("Мониторинг: бағалау рубрикасы", { placeholder: "title" });
  const hd2 = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: H.dk2 }, align: "center" } });
  const rub = [
    [hd2("Критерий"), hd2("1 – Бастапқы"), hd2("2 – Жеткілікті"), hd2("3 – Жоғары")],
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

  // ================= 19. CONCLUSION =================
  s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Мониторинг және қорытынды" });
  txt(s, "Қорытынды: пән мақсатына қалай жетеміз", { x: 0.7, y: 0.6, w: 11.9, h: 0.8, fontSize: 32, bold: true, color: C.background1, fontFace: THEME.headFontFace });
  const concl = [
    [I.layers, "Құзыреттілік тәсіл", "білім, білік, құндылық және рефлексияны бір тапсырмада біріктіреді (1-міндет)"],
    [I.chart, "Деңгейлік мазмұн", "Слэйтер ережесі базалықтан жоғары деңгейге дейін саралауға мүмкіндік береді (2-міндет)"],
    [I.robot, "ЖИ + проблемалық оқыту", "ChatGPT пен Claude дайын жауап көзі емес, тексерілетін серіктес болады (3-міндет)"],
    [I.check, "Мониторинг", "Рубрика, тест және ЖИ-диалог портфолиосы нәтижені объективті көрсетеді (4-міндет)"],
  ];
  concl.forEach((c, i) => {
    const y = 1.75 + i * 1.12;
    circleIcon(s, c[0], 0.7, y, 0.75, C.accent1);
    txt(s, [{ text: c[1] + ": ", options: { bold: true, color: C.accent2 } }, { text: c[2] }], { x: 1.7, y: y - 0.08, w: 7.7, h: 0.95, fontSize: 15, color: C.background1, valign: "middle" });
  });
  txt(s, [
    { text: "Диссертациямен байланысы: ", options: { bold: true, color: C.accent2 } },
    { text: "бұл тәсіл «Бейорганикалық химияны проблемалық оқытуда ЖИ құралдарын кіріктіру арқылы студенттердің сыни ойлауын дамыту» тақырыбындағы зерттеуімнің әдістемелік негізі." },
  ], { x: 0.7, y: 6.25, w: 8.7, h: 0.8, fontSize: 13, color: C.background2, valign: "middle" });
  atomRings(s, 11.1, 3.6, 1.6, "FFFFFF", 40);
  txt(s, "Назарларыңызға рахмет!", { x: 9.4, y: 6.05, w: 3.5, h: 0.6, fontSize: 18, bold: true, color: C.accent2, align: "center", valign: "middle" });
  s.addNotes("Қорытындылай келе, пәннің төрт міндетіне қалай жауап бергенімізді көрсетейін. Құзыреттілік тәсіл білім мен білікті ғана емес, құндылық пен рефлексияны да бір тапсырмада біріктіреді. Слэйтер ережесі мазмұнды базалықтан жоғары деңгейге дейін саралауға мүмкіндік береді. ЖИ проблемалық оқытумен бірге қолданылғанда дайын жауап көзі емес, тексерілетін серіктес болады. Ал рубрика мен портфолио мониторингті объективті етеді. Нәтижесінде пәннің басты мақсаты, яғни оқушыны ойлануға, өз бетінше шешім қабылдауға және шешімі үшін жауап беруге үйрету, нақты сабақ арқылы жүзеге асады. Бұл тәсіл менің диссертациямның әдістемелік негізі болып табылады. Назарларыңызға рахмет! (≈2,5 мин + сұрақтар)");

  // ================= 20. REFERENCES =================
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Мониторинг және қорытынды" });
  s.addText("Пайдаланылған әдебиеттер", { placeholder: "title" });
  const refs = [
    "Қазақстан Республикасы Ғылым және жоғары білім министрінің 2022 жылғы 20 шілдедегі № 2 бұйрығы «Жоғары және жоғары оқу орнынан кейінгі білім берудің мемлекеттік жалпыға міндетті стандарттарын бекіту туралы».",
    "CHBBKTJI 6305 «Химиялық білім берудегі құзыреттілік тәсіл және жасанды интеллект» пәнінің силлабусы. – Қызылорда: Қорқыт Ата атындағы ҚУ, 2026.",
    "Хуторской А.В. Ключевые компетенции как компонент личностно-ориентированной парадигмы образования // Народное образование. – 2003. – № 2. – С. 58–64.",
    "Зимняя И.А. Ключевые компетенции – новая парадигма результата образования // Высшее образование сегодня. – 2003. – № 5. – С. 34–42.",
    "Gonzalez J., Wagenaar R. (Eds.) Tuning Educational Structures in Europe. Final Report. Phase One. – Bilbao: University of Deusto, 2003.",
    "Anderson L.W., Krathwohl D.R. (Eds.) A Taxonomy for Learning, Teaching, and Assessing. – New York: Longman, 2001.",
    "Facione P.A. Critical Thinking: A Statement of Expert Consensus for Purposes of Educational Assessment and Instruction (The Delphi Report). – Millbrae, CA: California Academic Press, 1990.",
    "Махмутов М.И. Проблемное обучение: основные вопросы теории. – М.: Педагогика, 1975.",
    "Лернер И.Я. Дидактические основы методов обучения. – М.: Педагогика, 1981.",
    "Slater J.C. Atomic Shielding Constants // Physical Review. – 1930. – Vol. 36, № 1. – P. 57–64. – DOI: 10.1103/PhysRev.36.57",
    "Clementi E., Raimondi D.L. Atomic Screening Constants from SCF Functions // J. Chem. Phys. – 1963. – Vol. 38, № 11. – P. 2686–2689. – DOI: 10.1063/1.1733573",
    "Housecroft C.E., Sharpe A.G. Inorganic Chemistry. – 5th ed. – Harlow: Pearson, 2018.",
    "Clark T.M. Investigating the Use of an Artificial Intelligence Chatbot with General Chemistry Exam Questions // J. Chem. Educ. – 2023. – Vol. 100, № 5. – P. 1905–1916. – DOI: 10.1021/acs.jchemed.3c00027",
    "Kasneci E. et al. ChatGPT for good? On opportunities and challenges of large language models for education // Learning and Individual Differences. – 2023. – Vol. 103. – 102274. – DOI: 10.1016/j.lindif.2023.102274",
    "UNESCO. Guidance for generative AI in education and research. – Paris: UNESCO, 2023.",
    "Miao F., Cukurova M. AI competency framework for teachers. – Paris: UNESCO, 2024.",
  ];
  const half = 8;
  [refs.slice(0, half), refs.slice(half)].forEach((col, ci) => {
    txt(s, col.map((r, i) => ({ text: r, options: { bullet: { type: "number", startAt: ci * half + 1 }, breakLine: i < col.length - 1 } })),
      { x: 0.6 + ci * 6.15, y: 1.4, w: 5.95, h: 5.5, fontSize: 10.5, paraSpaceAfter: 5 });
  });
  s.addNotes("Әдебиеттер үш топқа бөлінеді: нормативтік құжаттар мен силлабус; құзыреттілік тәсіл мен педагогика (Хуторской, Зимняя, TUNING, Блум таксономиясы, Facione, Махмутов, Лернер); химиялық негіз (Слэйтер, Клементи–Раймонди, Housecroft) және ЖИ бойынша заманауи зерттеулер мен ЮНЕСКО құжаттары.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
