const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "VR Kasutajakogemus";
pres.author = "Martin";

// Color palette
const C = {
  bg: "0A0F1E",         // deep navy
  bgCard: "131B2E",     // slightly lighter card bg
  purple: "7C3AED",     // vibrant purple
  purpleLight: "A78BFA",
  cyan: "06B6D4",       // vibrant cyan
  cyanLight: "67E8F9",
  white: "FFFFFF",
  muted: "94A3B8",
  accent: "F472B6",     // pink accent for variety
};

const makeShadow = () => ({ type: "outer", blur: 12, offset: 3, angle: 135, color: "000000", opacity: 0.35 });

// ─── SLIDE 1: TITLE / HOOK ───────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  // Top glow blob (purple, large oval)
  s.addShape(pres.shapes.OVAL, {
    x: -1.5, y: -1.5, w: 6, h: 4,
    fill: { color: C.purple, transparency: 82 },
    line: { color: C.purple, transparency: 100 },
  });

  // Bottom-right glow blob (cyan)
  s.addShape(pres.shapes.OVAL, {
    x: 6.5, y: 3, w: 5, h: 4,
    fill: { color: C.cyan, transparency: 85 },
    line: { color: C.cyan, transparency: 100 },
  });

  // Thin cyan accent bar on left
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.55, y: 1.5, w: 0.06, h: 2.6,
    fill: { color: C.cyan },
    line: { color: C.cyan, transparency: 100 },
  });

  // Hook quote
  s.addText('"Esimest korda panin pähe Meta Quest 2\nja unusin, et sein on seal."', {
    x: 0.8, y: 1.5, w: 7.5, h: 1.3,
    fontSize: 16, color: C.muted, italic: true,
    fontFace: "Calibri", align: "left", valign: "top", margin: 0,
  });

  // Main title
  s.addText("VR Kasutajakogemus", {
    x: 0.8, y: 2.9, w: 8, h: 1.1,
    fontSize: 44, bold: true, color: C.white,
    fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
  });

  // Subtitle
  s.addText("7 aastat kogemust · 2 peaplaati · disaineri pilguga", {
    x: 0.8, y: 4.05, w: 8, h: 0.55,
    fontSize: 14, color: C.cyanLight, fontFace: "Calibri",
    align: "left", valign: "middle", margin: 0,
  });

  // Speaker notes
  s.addNotes("Alusta isikliku looga — esimene kord VR-is. Loo emotsionaalne seos. Tutvusta ennast lühidalt: 7 aastat kogemust, 2 peaplaati. Ütle, mis täna jutuks tuleb.");
}

// ─── SLIDE 2: MIKS VR UX ON TEISTSUGUNE ──────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  // Purple glow top-right
  s.addShape(pres.shapes.OVAL, {
    x: 7, y: -1, w: 5, h: 4,
    fill: { color: C.purple, transparency: 85 },
    line: { color: C.purple, transparency: 100 },
  });

  // Slide label
  s.addText("01  —  KONTEKST", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: "Calibri",
    align: "left", valign: "middle", margin: 0, charSpacing: 3,
  });

  // Title
  s.addText("Miks VR UX on teistsugune?", {
    x: 0.5, y: 0.72, w: 9, h: 0.72,
    fontSize: 32, bold: true, color: C.white,
    fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
  });

  // 4 cards in a 2x2 grid
  const cards = [
    { title: "Kehalisus", body: "Sina OLED sisu sees — mitte vaatleja. Keha asend ja liikumine loevad.", x: 0.5, y: 1.65, accent: C.purple },
    { title: "Ruumiline UI", body: "Nupud, menüüd ja tekstid on 3D-ruumis. Sügavus ja kaugus mõjutavad loetavust.", x: 5.2, y: 1.65, accent: C.cyan },
    { title: "Puudub klaviatuur & hiir", body: "Kontroller, käsijälgimine või pilk — kõik tuleb ümber mõelda.", x: 0.5, y: 3.4, accent: C.cyan },
    { title: "Kohalolek kui eesmärk", body: '"Presence" on VR-is edu mõõdupuu — hea UX ei katkesta seda tunnet.', x: 5.2, y: 3.4, accent: C.purple },
  ];

  cards.forEach(c => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: c.x, y: c.y, w: 4.4, h: 1.55,
      fill: { color: C.bgCard }, shadow: makeShadow(),
      line: { color: C.bgCard, transparency: 100 },
    });
    // Left accent bar
    s.addShape(pres.shapes.RECTANGLE, {
      x: c.x, y: c.y, w: 0.055, h: 1.55,
      fill: { color: c.accent },
      line: { color: c.accent, transparency: 100 },
    });
    s.addText(c.title, {
      x: c.x + 0.18, y: c.y + 0.1, w: 4.1, h: 0.38,
      fontSize: 14, bold: true, color: C.white,
      fontFace: "Calibri", align: "left", valign: "top", margin: 0,
    });
    s.addText(c.body, {
      x: c.x + 0.18, y: c.y + 0.52, w: 4.1, h: 0.9,
      fontSize: 12, color: C.muted,
      fontFace: "Calibri", align: "left", valign: "top", margin: 0,
    });
  });

  s.addNotes("Selgita, miks VR on disainerile täiesti uus väljakutse — ei saa eeldusi tuua 2D veebist. Tõsta esile kohalolek: kui UX on halb, see katkeb ja kogemus kaob.");
}

// ─── SLIDE 3: VÄLJAKUTSED (1/2) ───────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  s.addShape(pres.shapes.OVAL, {
    x: -2, y: 2, w: 5, h: 5,
    fill: { color: C.cyan, transparency: 88 },
    line: { color: C.cyan, transparency: 100 },
  });

  s.addText("02  —  VÄLJAKUTSED", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: "Calibri",
    align: "left", valign: "middle", margin: 0, charSpacing: 3,
  });

  s.addText("Peamised UX väljakutsed", {
    x: 0.5, y: 0.72, w: 9, h: 0.72,
    fontSize: 32, bold: true, color: C.white,
    fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
  });

  // Challenge rows (left side)
  const challenges = [
    {
      num: "01",
      title: "Sissejuhatus & õppimiskõver",
      body: "Uus kasutaja ei tea, kuhu vaadata, kuidas liikuda ega mida vajutada. Esimene 5 minutit on kõige kriitilisem.",
      personal: "Olen näinud, kuidas inimesed paanikas peaplaadi maha võtavad.",
    },
    {
      num: "02",
      title: "Liikumishaigus & mugavus",
      body: "Vale liikumise disain põhjustab iiveldust. Teleport vs vaba liikumine — see on disaineritele raske valik.",
      personal: "Esimestel aastatel piirasin sessioone 20 minutiga.",
    },
  ];

  challenges.forEach((c, i) => {
    const y = 1.7 + i * 1.75;
    // Number badge
    s.addShape(pres.shapes.OVAL, {
      x: 0.5, y: y, w: 0.52, h: 0.52,
      fill: { color: C.purple },
      line: { color: C.purple, transparency: 100 },
    });
    s.addText(c.num, {
      x: 0.5, y: y, w: 0.52, h: 0.52,
      fontSize: 11, bold: true, color: C.white,
      fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
    });
    s.addText(c.title, {
      x: 1.2, y: y, w: 8.3, h: 0.42,
      fontSize: 16, bold: true, color: C.white,
      fontFace: "Calibri", align: "left", valign: "top", margin: 0,
    });
    s.addText(c.body, {
      x: 1.2, y: y + 0.44, w: 8.3, h: 0.52,
      fontSize: 12, color: C.muted,
      fontFace: "Calibri", align: "left", valign: "top", margin: 0,
    });
    // Personal experience tag
    s.addText("💬 " + c.personal, {
      x: 1.2, y: y + 0.88, w: 8.3, h: 0.45,
      fontSize: 12, color: C.accent, italic: true,
      fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
    });
  });

  s.addNotes("Räägi oma kogemustest mõlema punktiga. Onboarding: mida hea disain teeks teisiti? Liikumishaigus: Half-Life: Alyx on hea näide — teleport + vaba liikumine koos, kasutaja valib.");
}

// ─── SLIDE 4: VÄLJAKUTSED (2/2) ───────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  s.addShape(pres.shapes.OVAL, {
    x: 7, y: 2, w: 5, h: 5,
    fill: { color: C.purple, transparency: 88 },
    line: { color: C.purple, transparency: 100 },
  });

  s.addText("02  —  VÄLJAKUTSED", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: "Calibri",
    align: "left", valign: "middle", margin: 0, charSpacing: 3,
  });

  s.addText("Peamised UX väljakutsed", {
    x: 0.5, y: 0.72, w: 9, h: 0.72,
    fontSize: 32, bold: true, color: C.white,
    fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
  });

  const challenges = [
    {
      num: "03",
      title: "Sisendite disain",
      body: "Kontrollerid, käsijälgimine, pilkjuhtimine — igaüks vajab erinevat interaktsioonimudelit. Ühe lahendusega ei tule kaugele.",
      personal: "Käsijälgimine Quest 3-s: intuitiivne, aga väsitav pärast 20 minutit.",
    },
    {
      num: "04",
      title: "UI paigutus 3D ruumis",
      body: "Kus menüü asub? Kui kaugel? Mis nurga all? 2D reeglid ei tööta. HUD, maailmas ankurdatud UI ja pihuarvuti stiil — kõik erinevad.",
      personal: "Halvad menüüd tekitavad kaelavalusid — kogemus räägib.",
    },
  ];

  challenges.forEach((c, i) => {
    const y = 1.7 + i * 1.75;
    s.addShape(pres.shapes.OVAL, {
      x: 0.5, y: y, w: 0.52, h: 0.52,
      fill: { color: C.cyan },
      line: { color: C.cyan, transparency: 100 },
    });
    s.addText(c.num, {
      x: 0.5, y: y, w: 0.52, h: 0.52,
      fontSize: 11, bold: true, color: C.white,
      fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
    });
    s.addText(c.title, {
      x: 1.2, y: y, w: 8.3, h: 0.42,
      fontSize: 16, bold: true, color: C.white,
      fontFace: "Calibri", align: "left", valign: "top", margin: 0,
    });
    s.addText(c.body, {
      x: 1.2, y: y + 0.44, w: 8.3, h: 0.52,
      fontSize: 12, color: C.muted,
      fontFace: "Calibri", align: "left", valign: "top", margin: 0,
    });
    s.addText("💬 " + c.personal, {
      x: 1.2, y: y + 0.88, w: 8.3, h: 0.45,
      fontSize: 12, color: C.accent, italic: true,
      fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
    });
  });

  s.addNotes("Sisend: too konkreetsed näited — Beat Saber vs Gravity Sketch. UI ruumis: ankurdatud käele on mugav, aga katkestab kohaloleku. Räägi isiklikust kogemusest kaelavaludega.");
}

// ─── SLIDE 5: HEA VR UX ───────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  s.addShape(pres.shapes.OVAL, {
    x: 3, y: -2, w: 6, h: 5,
    fill: { color: C.cyan, transparency: 87 },
    line: { color: C.cyan, transparency: 100 },
  });

  s.addText("03  —  NÄITED", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: "Calibri",
    align: "left", valign: "middle", margin: 0, charSpacing: 3,
  });

  s.addText("Mis näeb välja hea VR UX?", {
    x: 0.5, y: 0.72, w: 9, h: 0.72,
    fontSize: 32, bold: true, color: C.white,
    fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
  });

  // Two example cards side by side
  const examples = [
    {
      game: "Half-Life: Alyx",
      label: "Mängimine",
      points: [
        "Maailmas ankurdatud inventar — ei katkesta kohalolekut",
        "Liikumisvalikud: teleport + vaba — kasutaja otsustab",
        "Käsijälgimine enne, kui see oli standard",
        "Tutorial peidetud mängimise sisse",
      ],
      accent: C.purple,
      x: 0.5,
    },
    {
      game: "Gravity Sketch",
      label: "Tootlikkus",
      points: [
        "3D modelleerimine tundub loomulik, mitte teisendus 2D-st",
        "Kontekstuaalsed tööriistad ilmuvad käe lähedal",
        "Minimaalne menüü, maksimaalne tööpind",
        "Käsijälgimine vähendab kontrollerist sõltuvust",
      ],
      accent: C.cyan,
      x: 5.2,
    },
  ];

  examples.forEach(ex => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: ex.x, y: 1.65, w: 4.4, h: 3.65,
      fill: { color: C.bgCard }, shadow: makeShadow(),
      line: { color: C.bgCard, transparency: 100 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: ex.x, y: 1.65, w: 4.4, h: 0.08,
      fill: { color: ex.accent },
      line: { color: ex.accent, transparency: 100 },
    });
    // Label badge
    s.addShape(pres.shapes.RECTANGLE, {
      x: ex.x + 0.18, y: 1.82, w: 1.1, h: 0.28,
      fill: { color: ex.accent, transparency: 75 },
      line: { color: ex.accent, transparency: 100 },
    });
    s.addText(ex.label.toUpperCase(), {
      x: ex.x + 0.18, y: 1.82, w: 1.1, h: 0.28,
      fontSize: 8, bold: true, color: C.white,
      fontFace: "Calibri", align: "center", valign: "middle", margin: 0, charSpacing: 2,
    });
    s.addText(ex.game, {
      x: ex.x + 0.18, y: 2.18, w: 4.0, h: 0.45,
      fontSize: 18, bold: true, color: C.white,
      fontFace: "Calibri", align: "left", valign: "top", margin: 0,
    });
    s.addText(ex.points.map(p => ({ text: p, options: { bullet: true, breakLine: true, paraSpaceAfter: 4 } })), {
      x: ex.x + 0.18, y: 2.72, w: 4.0, h: 2.4,
      fontSize: 11.5, color: C.muted,
      fontFace: "Calibri", align: "left", valign: "top", margin: 0,
    });
  });

  s.addNotes("Räägi konkreetsete näidete kaudu. Mis disainiotsused muutsid kogemuse eriliseks? Mida saab disainer siit õppida? Half-Life: Alyx on kõigi aegade parim VR UX benchmark.");
}

// ─── SLIDE 6: TULEVIK ─────────────────────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  s.addShape(pres.shapes.OVAL, {
    x: 5, y: 1, w: 7, h: 7,
    fill: { color: C.purple, transparency: 88 },
    line: { color: C.purple, transparency: 100 },
  });

  s.addText("04  —  TULEVIK", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: "Calibri",
    align: "left", valign: "middle", margin: 0, charSpacing: 3,
  });

  s.addText("Kuhu see liigub?", {
    x: 0.5, y: 0.72, w: 7, h: 0.72,
    fontSize: 32, bold: true, color: C.white,
    fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
  });

  const trends = [
    { icon: "✋", title: "Käsijälgimine", body: "Kontrollerid kaovad. Käed on loomulikum ja intuitiivsem sisend — Quest 3 näitab suunda." },
    { icon: "🌐", title: "Segareaalsus (MR)", body: "VR + füüsiline maailm ühes. Apple Vision Pro ja Meta Quest 3 toovad MR disaineritele uued väljakutsed." },
    { icon: "🤖", title: "AI-juhitavad liidesed", body: "Häälassistendid, kontekstitundlikud menüüd ja adaptiivsed kogemused — AI muudab VR UX-i personaalsemaks." },
  ];

  trends.forEach((t, i) => {
    const y = 1.7 + i * 1.2;
    s.addText(t.icon, {
      x: 0.5, y: y, w: 0.6, h: 0.6,
      fontSize: 22, align: "center", valign: "middle", margin: 0,
    });
    s.addText(t.title, {
      x: 1.25, y: y, w: 7.5, h: 0.35,
      fontSize: 15, bold: true, color: C.white,
      fontFace: "Calibri", align: "left", valign: "top", margin: 0,
    });
    s.addText(t.body, {
      x: 1.25, y: y + 0.37, w: 7.5, h: 0.55,
      fontSize: 12, color: C.muted,
      fontFace: "Calibri", align: "left", valign: "top", margin: 0,
    });
    if (i < trends.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: 0.5, y: y + 1.02, w: 8.8, h: 0,
        line: { color: "1E293B", width: 1 },
      });
    }
  });

  s.addNotes("Hoia see lühike — üks lause trendi kohta. Rõhuta, et disaineritel on nüüd võimalus kujundada, kuidas inimesed neid tehnoloogiaid kasutavad — see on põnev aeg.");
}

// ─── SLIDE 7: KOKKUVÕTE + KÜSIMUSED ──────────────────────────────────────────
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  // Full background glow
  s.addShape(pres.shapes.OVAL, {
    x: -2, y: -2, w: 8, h: 8,
    fill: { color: C.purple, transparency: 90 },
    line: { color: C.purple, transparency: 100 },
  });
  s.addShape(pres.shapes.OVAL, {
    x: 6, y: 1, w: 6, h: 6,
    fill: { color: C.cyan, transparency: 90 },
    line: { color: C.cyan, transparency: 100 },
  });

  s.addText("05  —  KOKKUVÕTE", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: "Calibri",
    align: "left", valign: "middle", margin: 0, charSpacing: 3,
  });

  s.addText("3 asja, mida meeles pidada", {
    x: 0.5, y: 0.72, w: 9, h: 0.72,
    fontSize: 32, bold: true, color: C.white,
    fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
  });

  const takeaways = [
    { num: "1", text: "VR UX algab kehast, mitte ekraanist — disaini kehalisust, mitte klikke." },
    { num: "2", text: "Kohalolek on eesmärk — iga disainiotsus kas toetab või katkestab selle." },
    { num: "3", text: "Proovi ise — VR UX-i ei saa mõista ainult lugemisest. Tunded annavad andmed." },
  ];

  takeaways.forEach((t, i) => {
    const y = 1.75 + i * 1.05;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 9, h: 0.85,
      fill: { color: C.bgCard }, shadow: makeShadow(),
      line: { color: C.bgCard, transparency: 100 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: y, w: 0.055, h: 0.85,
      fill: { color: i === 1 ? C.purple : C.cyan },
      line: { color: i === 1 ? C.purple : C.cyan, transparency: 100 },
    });
    s.addText(t.num, {
      x: 0.7, y: y, w: 0.5, h: 0.85,
      fontSize: 22, bold: true, color: i === 1 ? C.purpleLight : C.cyanLight,
      fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
    });
    s.addText(t.text, {
      x: 1.25, y: y, w: 7.9, h: 0.85,
      fontSize: 13.5, color: C.white,
      fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
    });
  });

  // Questions line
  s.addShape(pres.shapes.LINE, {
    x: 0.5, y: 5.0, w: 9, h: 0,
    line: { color: "1E293B", width: 1 },
  });
  s.addText("Küsimused? 🎧", {
    x: 0.5, y: 5.1, w: 9, h: 0.4,
    fontSize: 18, bold: true, color: C.cyanLight,
    fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
  });

  s.addNotes("Võta iga punkt lühidalt kokku oma sõnadega. Lõpeta sõnumiga: 'parim viis VR UX-i mõista on seda kogeda.' Kutsu küsimusi, ole avatud arutelule.");
}

// ─── SLIDE TRANSITIONS ────────────────────────────────────────────────────────
// Apply transitions to all slides after creation
pres.slides.forEach((slide, i) => {
  const transitions = [
    { type: "fade",   dur: 1000 },  // Slide 1: title
    { type: "push",   dir: "u", dur: 700 },  // Slide 2
    { type: "push",   dir: "u", dur: 700 },  // Slide 3
    { type: "push",   dir: "u", dur: 700 },  // Slide 4
    { type: "reveal", dir: "l", dur: 700 },  // Slide 5
    { type: "push",   dir: "u", dur: 700 },  // Slide 6
    { type: "fade",   dur: 1200 },  // Slide 7: outro
  ];
  if (transitions[i]) slide.transition = transitions[i];
});

// ─── WRITE FILE ───────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "C:\\Users\\Martin\\Downloads\\ux\\vr_ux_presentation.pptx" })
  .then(() => console.log("✅  Salvestatud: vr_ux_presentation.pptx"))
  .catch(err => { console.error("❌  Viga:", err); process.exit(1); });
