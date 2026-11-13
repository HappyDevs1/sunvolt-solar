// Builds report/SunVolt_Presentation.pptx (Question 4).
// Needs pptxgenjs, react, react-dom, react-icons and sharp:  npm install pptxgenjs react react-dom react-icons sharp
// Run from the project root:  node analysis/build_deck.js && python analysis/add_deck_charts.py
// Charts are added by add_deck_charts.py (python-pptx) so they also display in Keynote.
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaMapMarkedAlt, FaTruck, FaUsers, FaUserCheck } = require("react-icons/fa");

const OUT = path.join(__dirname, "..", "report", "SunVolt_Presentation.pptx");
const APPLY_THEME = process.env.APPLY_THEME_JS; // path to the pptx skill's apply_theme.js (optional)

const HEX = {
  ink: "0B0B0B", white: "FFFFFF", navy: "14304F", paper: "F4F3EF",
  blue: "2A78D6", orange: "EB6834", grey: "C9C8C2", ink2: "52514E", ice: "D7E3F2",
};
const THEME = {
  name: "SunVolt",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.paper,
    accent1: HEX.blue, accent2: HEX.orange, accent3: HEX.grey, accent4: HEX.ink2,
    accent5: "1BAF7A", accent6: "EDA100", hlink: HEX.blue, folHlink: "6250D6",
  },
};

async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "SunVolt Energy: Closing the Solar Sales Gap";
  const C = pres.SchemeColor;
  const W = 13.333;

  // ---------------------------------------------------------------- layouts
  pres.defineSlideMaster({
    title: "TITLE_DARK",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.0, w: 11.7, h: 1.9, fontSize: 40, bold: true, color: C.background1, valign: "bottom", align: "left" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.1, w: 11.7, h: 1.2, fontSize: 20, color: C.background2, valign: "top", align: "left" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 1.15, fontSize: 30, bold: true, color: C.text2, valign: "top", align: "left" }, text: "" } },
      { text: { text: "SunVolt Energy  |  Customer analysis of 1,000 customers", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent4, margin: 0 } } },
    ],
    slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: C.accent4, align: "right" },
  });
  pres.defineSlideMaster({
    title: "CLOSE_DARK",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.45, w: 12.1, h: 1.0, fontSize: 36, bold: true, color: C.background1, valign: "top", align: "left" }, text: "" } },
    ],
    slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: C.background2, align: "right" },
  });


  function card(slide, x, y, w, h, fill, name) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: "none" }, objectName: name });
  }
  function txt(slide, text, opts) {
    slide.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, fontSize: 16, valign: "top" }, opts));
  }

  // ---------------------------------------------------------------- 1. Title
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Opening" });
  s.addText("Closing the Gap: Where SunVolt Can Win More Solar Backup Sales", { placeholder: "title" });
  s.addText("Customer, regional and installer analysis for the CEO and investors", { placeholder: "body" });
  txt(s, "Presented by: ____________________", { x: 0.8, y: 6.3, w: 8, h: 0.4, fontSize: 14, color: C.background2, objectName: "Presenter" });
  s.addNotes(
    "Good morning everyone. Today I want to show you where SunVolt Energy is winning and losing solar backup sales, and what we can do about it. " +
    "Everything I show comes from our customer data: 1,000 customers across three regions, after removing duplicate records."
  );

  // ---------------------------------------------------------------- 2. Setting
  pres.addSection({ title: "The situation" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The situation" });
  s.addText("Demand for solar backup is there, but fewer than half of our customers buy", { placeholder: "title" });
  const tiles = [
    ["1,000", "customers analysed across North America, Europe and the Pacific", HEX.paper, C.text2],
    ["481", "of them bought a home solar backup system", HEX.paper, C.text2],
    ["48%", "overall purchase rate, so 519 customers who know us did not buy", "FDEBE3", C.accent2],
  ];
  tiles.forEach(([big, lab, fill, col], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.9, 3.8, 3.3, fill, "KPI tile " + (i + 1));
    txt(s, big, { x: x + 0.35, y: 2.2, w: 3.1, h: 1.3, fontSize: 66, bold: true, color: col, valign: "middle", objectName: "KPI value " + (i + 1) });
    txt(s, lab, { x: x + 0.35, y: 3.6, w: 3.1, h: 1.3, fontSize: 16, color: C.text1, objectName: "KPI label " + (i + 1) });
  });
  txt(s, [
    { text: "The opportunity: ", options: { bold: true, color: C.text2 } },
    { text: "our Pacific customers already buy at 59%. If we can get the rest of the business closer to that, we grow sales from customers we already have." },
  ], { x: 0.6, y: 5.6, w: 12.1, h: 0.9, fontSize: 18, objectName: "Opportunity line" });
  s.addNotes(
    "Let me start with the big picture. We looked at 1,000 customers. 481 of them bought a solar backup system, which is a purchase rate of 48 percent. " +
    "Electricity costs keep rising and households want reliable power, so the demand is clearly there. But it also means more than half of the people who already know SunVolt walked away without buying. " +
    "The good news is that one part of our business already shows what is possible: in the Pacific, 59 percent of customers buy. The question for today is why the rest of the business is behind, and how we close that gap."
  );

  // ---------------------------------------------------------------- 3. Region
  pres.addSection({ title: "Where and why" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Where and why" });
  s.addText("Our biggest market, North America, converts the worst", { placeholder: "title" });
  card(s, 8.7, 1.9, 4.0, 4.5, HEX.paper, "North America callout");
  txt(s, "508", { x: 9.05, y: 2.15, w: 3.3, h: 1.1, fontSize: 60, bold: true, color: C.accent2, objectName: "NA customers value" });
  txt(s, "customers live in North America, half of our entire customer base", { x: 9.05, y: 3.3, w: 3.3, h: 1.0, fontSize: 16 });
  txt(s, [
    { text: "Only 43% buy, ", options: { bold: true } },
    { text: "compared with 59% in the Pacific. Matching the Pacific would add about 79 sales." },
  ], { x: 9.05, y: 4.5, w: 3.3, h: 1.6, fontSize: 16, objectName: "NA gap text" });
  s.addNotes(
    "When we split the customers by region, the picture changes. In the Pacific, 59 percent of customers buy. Europe is close to our average at 49 percent. North America is the lowest at only 43 percent. " +
    "That would not matter much if North America were small, but it is our biggest market. 508 customers, half of everyone we serve, live there. So our biggest region is also our weakest one. " +
    "If North America converted like the Pacific does, we would make about 79 more sales from customers we already have."
  );

  // ---------------------------------------------------------------- 4. Distance
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Where and why" });
  s.addText("Customers more than 5 km from an installer rarely buy, especially in North America", { placeholder: "title" });
  card(s, 8.7, 1.9, 4.0, 4.5, HEX.paper, "Distance callout");
  txt(s, "51% to 27%", { x: 9.05, y: 2.15, w: 3.4, h: 1.0, fontSize: 40, bold: true, color: C.accent2, objectName: "NA distance drop" });
  txt(s, "is how far buying falls in North America once a customer lives more than 5 km from an installer", { x: 9.05, y: 3.2, w: 3.3, h: 1.3, fontSize: 16 });
  txt(s, [
    { text: "171 North American customers ", options: { bold: true } },
    { text: "live beyond 5 km. That is one in three customers in the region." },
  ], { x: 9.05, y: 4.65, w: 3.3, h: 1.4, fontSize: 16, objectName: "NA distance count" });
  s.addNotes(
    "So why is North America behind? The first reason is distance to an installer. Across all customers, people who live within 5 kilometres of an installer buy at 53 percent. Beyond 5 kilometres, that falls to 36 percent. " +
    "This chart shows it by region. In every region the orange bar, customers beyond 5 kilometres, is lower. But the drop is sharpest in North America, from 51 percent down to 27 percent. " +
    "And this is not a small group. 171 North American customers live beyond 5 kilometres, about one in three customers in that region. In Europe it is only 34 people. " +
    "In simple terms, when installation is far away, it feels harder and more expensive, and customers do not go ahead."
  );

  // ---------------------------------------------------------------- 5. Married
  pres.addSection({ title: "Who" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Who" });
  s.addText("Married customers over 40 are the largest group not buying", { placeholder: "title" });
  card(s, 8.7, 1.9, 4.0, 2.05, "FDEBE3", "Married 40+ callout");
  txt(s, "37%", { x: 9.05, y: 2.1, w: 1.6, h: 0.9, fontSize: 44, bold: true, color: C.accent2, objectName: "Married 40+ rate" });
  txt(s, "of the 384 married customers aged 40+ buy", { x: 9.05, y: 3.0, w: 3.3, h: 0.8, fontSize: 16 });
  card(s, 8.7, 4.25, 4.0, 2.15, "FDEBE3", "Families callout");
  txt(s, "34%", { x: 9.05, y: 4.45, w: 1.6, h: 0.9, fontSize: 44, bold: true, color: C.accent2, objectName: "Families rate" });
  txt(s, "of married households with 2 or more dependents buy (314 customers)", { x: 9.05, y: 5.35, w: 3.3, h: 0.9, fontSize: 16 });
  s.addNotes(
    "The second reason is the type of household. Our strongest buyers are customers in their 30s. Both single and married customers in that age group buy at about 60 percent. " +
    "But look at what happens after 40. Single customers keep buying at more than 50 percent at every age. Married customers drop to about 40 percent, and to only 22 percent at 60 and over. " +
    "There are 384 married customers aged 40 and over, and only 37 percent of them buy. Married households with two or more dependents are even lower at 34 percent. Households with two or more vehicles show the same pattern. " +
    "These are the families with the most to lose in a power cut, so the need is there. The most likely barrier is cost, because a system big enough for a whole family is a bigger purchase. " +
    "It is also worth saying what does not matter: gender makes no difference, and income only matters below 30 thousand."
  );

  // ---------------------------------------------------------------- 6. Untapped
  pres.addSection({ title: "The cost" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "The cost" });
  s.addText("We are missing sales from customers we already have", { placeholder: "title" });
  card(s, 8.7, 1.9, 4.0, 4.5, HEX.paper, "North America upside callout");
  txt(s, "+79", { x: 9.05, y: 2.15, w: 3.3, h: 1.1, fontSize: 60, bold: true, color: C.accent1, objectName: "NA upside value" });
  txt(s, "extra sales if North America converted like the Pacific does", { x: 9.05, y: 3.3, w: 3.3, h: 1.0, fontSize: 16 });
  txt(s, "These customers already know SunVolt, so winning them costs far less than finding new ones. The groups overlap, so the bars should not be added together.", { x: 9.05, y: 4.5, w: 3.3, h: 1.8, fontSize: 14, color: C.accent4 });
  s.addNotes(
    "So what is this costing us? This chart shows how many extra sales we would make if each of these groups simply bought at our own company average of 48 percent. Not an ambitious target, just average. " +
    "Households with two or more vehicles would add about 56 sales. Married families with two or more dependents about 43. Married customers over 40 about 42. And North American customers beyond 5 kilometres about 35. " +
    "These groups overlap, so we should not add the numbers together. But the overall point is clear. If North America converted like the Pacific, that alone is about 79 more sales. " +
    "And because these are customers who already know us, these are some of the cheapest sales we can make, which is good for profitability as well as growth."
  );

  // ---------------------------------------------------------------- 7. Actions
  pres.addSection({ title: "What we do" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "What we do" });
  s.addText("Three actions to close the gap", { placeholder: "title" });
  const acts = [
    [FaMapMarkedAlt, "Bring installers closer", "ADOPTION", "Add partner installers or mobile install teams in North American areas more than 5 km from a branch.", "Target: 171 customers now buying at 27%"],
    [FaTruck, "Remove distance costs", "PROFITABILITY", "Free or flat-fee installation beyond 5 km, with installs grouped by area on set days to keep cost per job down.", "Beyond 5 km: 36% buy vs 53% within"],
    [FaUsers, "Launch a family bundle", "TARGETING", "A larger backup package on a monthly payment plan, marketed to married households over 40 and homes with 2+ vehicles.", "These groups buy at 34% to 37%"],
  ];
  for (let i = 0; i < acts.length; i++) {
    const [Ic, head, tag, body, proof] = acts[i];
    const x = 0.6 + i * 4.1;
    card(s, x, 1.7, 3.8, 4.35, HEX.paper, "Action card " + (i + 1));
    s.addShape(pres.shapes.OVAL, { x: x + 0.35, y: 2.0, w: 0.85, h: 0.85, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Icon circle " + (i + 1) });
    s.addImage({ data: await icon(Ic, HEX.white), x: x + 0.55, y: 2.2, w: 0.45, h: 0.45, altText: head });
    txt(s, tag, { x: x + 1.4, y: 2.25, w: 2.2, h: 0.35, fontSize: 12, bold: true, color: C.accent1, objectName: "Action tag " + (i + 1) });
    txt(s, head, { x: x + 0.35, y: 3.05, w: 3.2, h: 0.5, fontSize: 20, bold: true, color: C.text2, objectName: "Action heading " + (i + 1) });
    txt(s, body, { x: x + 0.35, y: 3.6, w: 3.2, h: 1.55, fontSize: 15, objectName: "Action body " + (i + 1) });
    txt(s, proof, { x: x + 0.35, y: 5.3, w: 3.2, h: 0.55, fontSize: 14, bold: true, color: C.accent2, objectName: "Action proof " + (i + 1) });
  }
  s.addImage({ data: await icon(FaUserCheck, HEX.blue), x: 0.6, y: 6.32, w: 0.35, h: 0.35, altText: "Core buyers" });
  txt(s, [
    { text: "Keep doing what works: ", options: { bold: true, color: C.text2 } },
    { text: "customers aged 30-39, professionals and Pacific customers already buy at 54% to 59%." },
  ], { x: 1.1, y: 6.3, w: 11.6, h: 0.45, fontSize: 15, valign: "middle", objectName: "Core buyers line" });
  s.addNotes(
    "So here is what I recommend. Three actions, each aimed at a gap we measured. " +
    "First, for adoption: bring installers closer. Add partner installers or mobile installation teams in the North American areas that are more than 5 kilometres from a branch. That targets 171 customers who currently buy at only 27 percent. " +
    "Second, for profitability: remove the cost of distance. Offer free or flat-fee installation to customers beyond 5 kilometres, and group installations in the same area on set days so our cost per installation stays low. " +
    "Third, for targeting: launch a family bundle. A larger backup package on a monthly payment plan, marketed directly to married households over 40 and homes with more than one vehicle. That also raises the value of each sale. " +
    "At the same time, we keep investing in what already works: customers in their 30s, professionals and the Pacific, who buy at 54 to 59 percent."
  );

  // ---------------------------------------------------------------- 8. Call to action
  pres.addSection({ title: "Decision" });
  s = pres.addSlide({ masterName: "CLOSE_DARK", sectionTitle: "Decision" });
  s.addText("Approve a 3-month North America pilot", { placeholder: "title" });
  txt(s, "48%", { x: 0.6, y: 1.8, w: 2.8, h: 1.3, fontSize: 72, bold: true, color: C.background2, objectName: "From rate" });
  txt(s, "today", { x: 0.6, y: 3.1, w: 2.8, h: 0.4, fontSize: 16, color: C.background2 });
  s.addShape(pres.shapes.RIGHT_ARROW, { x: 3.4, y: 2.2, w: 1.5, h: 0.6, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Arrow" });
  txt(s, "59%", { x: 5.2, y: 1.8, w: 2.8, h: 1.3, fontSize: 72, bold: true, color: C.background1, objectName: "To rate" });
  txt(s, "what the Pacific already achieves", { x: 5.2, y: 3.1, w: 3.2, h: 0.4, fontSize: 16, color: C.background2 });
  card(s, 8.8, 1.8, 3.9, 3.0, "1F4470", "Pilot measures card");
  txt(s, "We will measure", { x: 9.1, y: 2.0, w: 3.4, h: 0.4, fontSize: 16, bold: true, color: C.background1 });
  txt(s, [
    { text: "Purchase rate beyond 5 km (now 27%)", options: { bullet: true, breakLine: true } },
    { text: "Purchase rate of married customers (now 43%)", options: { bullet: true, breakLine: true } },
    { text: "Cost per installation", options: { bullet: true } },
  ], { x: 9.1, y: 2.5, w: 3.4, h: 2.1, fontSize: 15, color: C.background1, paraSpaceAfter: 6, objectName: "Pilot measures" });
  txt(s, "Big Idea", { x: 0.6, y: 5.05, w: 3, h: 0.4, fontSize: 14, bold: true, color: C.accent2 });
  txt(s, "If we bring installers closer to North American customers and build an offer for married families, we can lift our purchase rate from 48% towards the 59% the Pacific already shows is possible.",
    { x: 0.6, y: 5.45, w: 12.1, h: 1.2, fontSize: 20, color: C.background1, objectName: "Big Idea" });
  s.addNotes(
    "To close: we do not have a demand problem. We have a reach problem and an offer problem, and both can be fixed. " +
    "If we bring installers closer to our North American customers and build an offer for married families, we can lift our purchase rate from 48 percent towards the 59 percent the Pacific already shows is possible. " +
    "My request today is to approve a three-month pilot in North America. We will track three things: the purchase rate of customers beyond 5 kilometres, which is 27 percent today, the purchase rate of married customers, which is 43 percent today, and our cost per installation. " +
    "I will report back on the results at the end of the pilot so we can decide whether to roll it out to every region. Thank you, and I am happy to take any questions."
  );

  await pres.writeFile({ fileName: OUT });
  if (APPLY_THEME) {
    const { applyTheme } = require(APPLY_THEME);
    await applyTheme(OUT, THEME);
  }
  console.log("saved", OUT);
})();
