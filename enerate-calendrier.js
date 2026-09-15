const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, 
        ShadingType, PageNumber, LevelFormat, ExternalHyperlink } = require('docx');
const fs = require('fs');

// Border style
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// Helper for table cell
function cell(text, opts = {}) {
  const { bold = false, fill = null, width = 2340, align = AlignmentType.LEFT, fontSize = 18, color = "000000" } = opts;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text, bold, size: fontSize, font: "Arial", color })]
    })]
  });
}

function headerCell(text, width = 2340) {
  return cell(text, { bold: true, fill: "1F4E79", width, align: AlignmentType.CENTER, fontSize: 18, color: "FFFFFF" });
}

function sectionHeader(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, bold: true, size: 28, font: "Arial", color: "1F4E79" })]
  });
}

function subHeader(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 24, font: "Arial", color: "2E75B6" })]
  });
}

function subSubHeader(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text, bold: true, size: 20, font: "Arial", color: "404040" })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 20, font: "Arial", ...opts })]
  });
}

function boldP(text) {
  return new Paragraph({
    spacing: { after: 80 },
    children: [new TextRun({ text, bold: true, size: 20, font: "Arial" })]
  });
}

function bullet(text, ref = "bullets") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 60 },
    children: [new TextRun({ text, size: 20, font: "Arial" })]
  });
}

function linkP(label, url) {
  return new Paragraph({
    spacing: { after: 60 },
    children: [
      new ExternalHyperlink({
        children: [new TextRun({ text: label, style: "Hyperlink", size: 18, font: "Arial" })],
        link: url
      })
    ]
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "1F4E79" },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 20, bold: true, font: "Arial", color: "404040" },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "practice", levels: [{ level: 0, format: LevelFormat.BULLET, text: "→", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1008, right: 1008, bottom: 1008, left: 1008 } // 0.7 inch
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children: [new TextRun({ text: "Calendrier de formation révisé — Profil ONG / Développement opérationnel", size: 16, font: "Arial", color: "666666", italics: true })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page ", size: 16, font: "Arial", color: "666666" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "666666" }),
            new TextRun({ text: " | Version révisée – Septembre 2026 | Principe : Certification → Preuve d'application", size: 16, font: "Arial", color: "666666" })
          ]
        })]
      })
    },
    children: [
      // TITLE
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: "CALENDRIER DE FORMATION RÉVISÉ", bold: true, size: 36, font: "Arial", color: "1F4E79" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: "Vers un profil ONG / Développement opérationnel", bold: true, size: 26, font: "Arial", color: "2E75B6" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: "Version renforcée avec livrables précis, jalons de pratique et indicateurs de suivi", size: 18, font: "Arial", italics: true, color: "666666" })]
      }),

      // PROFIL
      sectionHeader("Profil de départ et objectifs"),
      p("Licence en sociologie du développement, Master en ingénierie juridique des collectivités territoriales décentralisées, titulaire de 3 certificats DisasterReady (action humanitaire, fondamentaux de la gestion de projet, MEAL)."),
      boldP("Objectif à 15 mois :"),
      bullet("Être compétitif pour des postes MEAL Officer / Project Officer / Assistant MEAL au Tchad ou en Afrique centrale."),
      bullet("Disposer d'un portfolio public de 5–7 livrables concrets."),
      bullet("Avoir au moins 3–4 mois d'expérience pratique cumulative (bénévolat + missions d'appui)."),
      bullet("Anglais opérationnel (rédaction de rapports courts + compréhension orale professionnelle)."),

      p("Durée totale : environ 15 mois (semaines 1 à 60)."),
      p("Principe central : chaque certification ou module doit être suivi d'une preuve d'application concrète documentée dans le portfolio."),

      // TABLEAU DE BORD
      sectionHeader("Tableau de bord de suivi (à tenir dès la semaine 1)"),
      p("Créer un fichier Excel ou Notion simple avec les colonnes suivantes et le mettre à jour chaque dimanche :"),
      
      new Table({
        width: { size: 9900, type: WidthType.DXA },
        columnWidths: [2200, 1800, 2000, 2000, 1900],
        rows: [
          new TableRow({ children: [
            headerCell("Indicateur", 2200),
            headerCell("Cible à 6 mois", 1800),
            headerCell("Cible à 12 mois", 2000),
            headerCell("Cible à 15 mois", 2000),
            headerCell("Suivi hebdo", 1900)
          ]}),
          new TableRow({ children: [
            cell("Certifications DPro obtenues", 2200),
            cell("Project + MEAL", 1800),
            cell("+ Finance", 2000),
            cell("3/3", 2000),
            cell("Oui / Non", 1900)
          ]}),
          new TableRow({ children: [
            cell("Livrables dans le portfolio", 2200),
            cell("3 minimum", 1800),
            cell("5 minimum", 2000),
            cell("6–7", 2000),
            cell("Nombre", 1900)
          ]}),
          new TableRow({ children: [
            cell("Heures de pratique terrain", 2200),
            cell("40 h", 1800),
            cell("100 h", 2000),
            cell("150 h+", 2000),
            cell("Cumul", 1900)
          ]}),
          new TableRow({ children: [
            cell("Anglais (min/jour moyen)", 2200),
            cell("40 min", 1800),
            cell("45–60 min", 2000),
            cell("45–60 min", 2000),
            cell("Moyenne", 1900)
          ]}),
          new TableRow({ children: [
            cell("Contacts réseau utiles", 2200),
            cell("10", 1800),
            cell("25", 2000),
            cell("40+", 2000),
            cell("Cumul", 1900)
          ]}),
          new TableRow({ children: [
            cell("Candidatures envoyées", 2200),
            cell("0–2", 1800),
            cell("5–8", 2000),
            cell("12+", 2000),
            cell("Cumul", 1900)
          ]}),
        ]
      }),

      new Paragraph({ spacing: { before: 160, after: 80 }, children: [
        new TextRun({ text: "Portfolio public recommandé : ", bold: true, size: 20, font: "Arial" }),
        new TextRun({ text: "Google Drive partagé (lien dans le CV/LinkedIn) ou dossier LinkedIn « Featured ».", size: 20, font: "Arial" })
      ]}),

      // PHASE 1
      sectionHeader("Phase 1 — Recrutabilité (Semaines 1 à 12)"),
      p("Objectif : obtenir les bases certifiantes et les outils quotidiens des ONG, tout en lançant l'anglais et le réseau."),

      subHeader("Semaines 1–4 : Project DPro Foundation"),
      bullet("Action : Relire le guide Project DPro + faire 2–3 examens blancs."),
      bullet("Ressource : Guide Project DPro – PM4NGOs (gratuit)."),
      linkP("https://pm4ngos.org/methodologies-guides/project-dpro/", "https://pm4ngos.org/methodologies-guides/project-dpro/"),
      bullet("Examen : 75 questions, livre ouvert, 2 h, seuil 65 %. Tarif Local : 30 USD."),
      linkP("Plateforme d'examens PM4NGOs", "https://pm4ngos.org/candidates-and-learners/exams-certificates/"),
      boldP("Livrable attendu (semaine 4) :"),
      bullet("Certificat Project DPro Foundation (PDF) ajouté au portfolio."),
      bullet("Fiche récapitulative personnelle d'1 page (principaux outils Project DPro + 3 leçons retenues)."),

      subHeader("Semaines 5–8 : KoboToolbox + CHS / PSEA"),
      bullet("KoboToolbox : autoformation via documentation officielle + communauté."),
      linkP("https://support.kobotoolbox.org/ et https://community.kobotoolbox.org/", "https://support.kobotoolbox.org/"),
      bullet("CHS (norme humanitaire fondamentale) + PSEA : modules DisasterReady (1 journée chacun)."),
      linkP("https://fr.disasterready.org/", "https://fr.disasterready.org/"),
      boldP("Livrables attendus (semaine 8) :"),
      bullet("Formulaire Kobo complet créé et testé (enquête de base ou suivi d'activités – 15–25 questions)."),
      bullet("Export CSV + analyse simple dans Excel (tableaux croisés)."),
      bullet("Certificats CHS + PSEA (PDF)."),
      bullet("Note d'1 page : « Comment j'appliquerais le CHS et le PSEA dans un projet terrain au Tchad »."),

      subHeader("Semaines 9–12 : Excel avancé + démarrage réseau"),
      bullet("Tableaux croisés dynamiques, RECHERCHEX / XLOOKUP, tableaux de suivi de projet, graphiques."),
      bullet("Ressources : YouTube francophone (Excel Formation, etc.) + OpenClassrooms si besoin."),
      boldP("Livrables attendus (semaine 12) :"),
      bullet("Fichier Excel modèle de suivi de projet (indicateurs, budget simple, chronogramme) prêt à être réutilisé."),
      bullet("Profil LinkedIn mis à jour (photo pro, résumé orienté résultats, compétences MEAL/Project)."),
      bullet("10 contacts utiles ajoutés (ONG locales Sarh/N'Djamena, clusters, alumni)."),

      subHeader("Fil rouge anglais (dès semaine 1)"),
      bullet("45 min/jour minimum (Duolingo pour la base + BBC Learning English + lecture d'articles ReliefWeb/OCHA)."),
      bullet("Production écrite : 1 résumé court (150–200 mots) en anglais par semaine à partir de la semaine 5."),
      boldP("Indicateur : moyenne hebdomadaire ≥ 40 min + au moins 8 résumés produits à la fin de Phase 1."),

      // PHASE 2
      sectionHeader("Phase 2 — Spécialisation MEAL + Première pratique (Semaines 13 à 26)"),
      p("Objectif : certification MEAL DPro + Power BI + première mission terrain documentée."),

      subHeader("Semaines 13–17 : MEAL DPro"),
      bullet("Relire le guide MEAL DPro + examens blancs."),
      linkP("https://pm4ngos.org/methodologies-guides/meal-dpro/", "https://pm4ngos.org/methodologies-guides/meal-dpro/"),
      bullet("Examen : même format, tarif Local 30 USD. Passage en semaine 17."),
      boldP("Livrable (semaine 17) : Certificat MEAL DPro + fiche 1 page des outils MEAL prioritaires."),

      subHeader("Semaines 18–24 : Power BI + préparation mission pratique"),
      bullet("Microsoft Learn – Power BI (parcours gratuit)."),
      linkP("https://learn.microsoft.com/fr-fr/training/powerplatform/power-bi", "https://learn.microsoft.com/fr-fr/training/powerplatform/power-bi"),
      boldP("Livrables (semaine 24) :"),
      bullet("Dashboard Power BI (ou Excel avancé si Power BI inaccessible) à partir d'un jeu de données fictif ou réel (indicateurs humanitaires ou développement)."),
      bullet("Capture d'écran + explication ½ page des visualisations choisies."),

      subHeader("Semaines 25–26 + extension pratique : Première mission terrain"),
      p("Chercher activement dès la semaine 18 une mission bénévole ou d'appui (association de Sarh, ONG nationale, projet local). Durée idéale : 3–6 semaines cumulées (pas forcément continues)."),
      boldP("Jalons de pratique obligatoires :"),
      bullet("Collecte de données réelle avec Kobo (ou papier + saisie)."),
      bullet("Nettoyage + analyse des données."),
      bullet("Rapport de suivi ou note de capitalisation (4–8 pages)."),
      boldP("Livrable final Phase 2 :"),
      bullet("Rapport de mission (anonymisé si nécessaire) + dataset + dashboard + certificat de reconnaissance de l'organisation d'accueil."),
      bullet("Ajout au portfolio + post LinkedIn (avec autorisation)."),

      // PHASE 3
      sectionHeader("Phase 3 — Rentabilité (Semaines 27 à 40)"),
      p("Objectif : compétences qui « rapportent » aux employeurs (propositions + finance de projet) + 2e cycle de pratique."),

      subHeader("Semaines 27–32 : Rédaction de propositions de financement"),
      bullet("Modules DisasterReady « Proposal Writing » + ressources Humentum (gratuits ou bas coût)."),
      boldP("Livrables :"),
      bullet("Note conceptuelle complète (3–5 pages) pour un projet fictif ou réel d'une association locale."),
      bullet("Cadre logique (logframe) associé."),
      bullet("Optionnel : réponse simulée à un appel à projets (ECHO, UE, fondation…)."),

      subHeader("Semaines 33–40 : Finance DPro"),
      bullet("Étudier le guide Finance DPro."),
      linkP("https://pm4ngos.org/methodologies-guides/finance-dpro/", "https://pm4ngos.org/methodologies-guides/finance-dpro/"),
      bullet("Examen : tarif Local 30 USD ou In-Country 85 USD selon statut. Passage semaine 40."),
      boldP("Livrable : Certificat Finance DPro + budget de projet simple commenté (Excel) lié à la note conceptuelle précédente."),

      subHeader("Pratique continue (semaines 27–40)"),
      bullet("Deuxième mission d'appui (idéalement plus responsabilisante : appui MEAL ou reporting)."),
      bullet("Objectif cumul : ≥ 100 heures de pratique terrain documentées."),

      // PHASE 4
      sectionHeader("Phase 4 — Différenciation & Positionnement (Semaines 41 à 60)"),
      p("Objectif : expertise gouvernance locale + anglais consolidé + candidatures actives + portfolio final."),

      subHeader("Semaines 41–48 : ALGA / CGLU Afrique"),
      bullet("Prioriser les modules gratuits de l'e-Academy (budget participatif, transformation digitale, soft skills…)."),
      linkP("https://www.uclga.org/alga/ et e-Academy CGLU Afrique", "https://www.uclga.org/alga/"),
      boldP("Livrable : 2–3 certificats de modules + note de synthèse (2 pages) sur « Comment l'expertise collectivités territoriales renforce mon profil MEAL/Project dans les ONG »."),

      subHeader("Semaines 49–54 : ONU-Habitat + UNITAR"),
      bullet("Cours gratuits sur gouvernance urbaine/territoriale et décentralisation."),
      linkP("https://urbanacademy.unhabitat.org/", "https://urbanacademy.unhabitat.org/"),
      linkP("https://www.unitar.org/", "https://www.unitar.org/"),
      boldP("Livrable : certificats de complétion + 1 article LinkedIn ou note de 1–2 pages reliant gouvernance locale et programmation ONG."),

      subHeader("Semaines 55–60 : Positionnement final et candidatures"),
      boldP("Actions prioritaires :"),
      bullet("Finaliser le portfolio (6–7 livrables minimum)."),
      bullet("CV + lettre de motivation orientés résultats (version FR + EN courte)."),
      bullet("LinkedIn optimisé + 3–5 posts techniques publiés sur la période."),
      bullet("Candidatures ciblées (minimum 8–12 sur les 4 derniers mois)."),
      bullet("Préparation entretiens (études de cas MEAL, questions comportementales)."),
      boldP("Indicateur de réussite Phase 4 : portfolio complet + au moins 8 candidatures envoyées + 40+ contacts réseau."),

      // RÉCAP COÛTS
      sectionHeader("Récapitulatif des coûts estimés (actualisé)"),
      new Table({
        width: { size: 9900, type: WidthType.DXA },
        columnWidths: [5500, 4400],
        rows: [
          new TableRow({ children: [
            headerCell("Poste", 5500),
            headerCell("Coût estimé", 4400)
          ]}),
          new TableRow({ children: [
            cell("Project DPro Foundation (examen – tarif Local)", 5500),
            cell("30 USD", 4400)
          ]}),
          new TableRow({ children: [
            cell("MEAL DPro (examen – tarif Local)", 5500),
            cell("30 USD", 4400)
          ]}),
          new TableRow({ children: [
            cell("Finance DPro (examen – tarif Local ou In-Country)", 5500),
            cell("30–85 USD", 4400)
          ]}),
          new TableRow({ children: [
            cell("Kobo, CHS, PSEA, Power BI, ONU-Habitat, UNITAR, Excel", 5500),
            cell("Gratuit", 4400)
          ]}),
          new TableRow({ children: [
            cell("Rédaction de propositions (DisasterReady / Humentum)", 5500),
            cell("Gratuit à faible coût", 4400)
          ]}),
          new TableRow({ children: [
            cell("ALGA e-Academy (modules gratuits prioritaires)", 5500),
            cell("Gratuit (certifications longues variables)", 4400)
          ]}),
          new TableRow({ children: [
            cell("Anglais (Duolingo, BBC, ressources libres)", 5500),
            cell("Gratuit", 4400)
          ]}),
          new TableRow({ children: [
            cell("Total estimé (hors formations ALGA longues)", 5500),
            cell("≈ 90–145 USD sur 15 mois", 4400)
          ]}),
        ]
      }),

      new Paragraph({ spacing: { before: 120, after: 80 }, children: [
        new TextRun({ text: "Important : ", bold: true, size: 18, font: "Arial" }),
        new TextRun({ text: "Toujours utiliser le sélecteur de tarif PM4NGOs. Les tarifs Local (30 USD) s'appliquent aux étudiants, volontaires, personnel d'ONG locales et résidents de pays en développement (Tchad inclus). Prévoir un éventuel audit de statut.", size: 18, font: "Arial" })
      ]}),

      // PRINCIPES
      sectionHeader("Les 5 principes non négociables de cette version révisée"),
      bullet("1. Certification → Preuve d'application documentée (portfolio)."),
      bullet("2. Pratique terrain cumulative ≥ 150 heures sur 15 mois (idéal)."),
      bullet("3. Anglais mesurable (temps + production écrite)."),
      bullet("4. Réseau et visibilité construits progressivement (pas seulement en fin de parcours)."),
      bullet("5. Candidatures démarrées avant la fin du plan (apprentissage par l'action)."),

      new Paragraph({ spacing: { before: 200 }, children: [
        new TextRun({ text: "Ce calendrier transforme un parcours de formation en un véritable projet professionnel mesurable. Le succès ne se mesure plus uniquement au nombre de certificats, mais à la qualité du portfolio et à la capacité à démontrer une expérience concrète.", size: 20, font: "Arial", italics: true })
      ]}),

      new Paragraph({ spacing: { before: 160 }, children: [
        new TextRun({ text: "Document généré en septembre 2026 – Version experte révisée.", size: 16, font: "Arial", color: "666666" })
      ]}),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("./Calendrier_Formation_ONG_Revise.docx", buffer);
  console.log("Document created successfully: Calendrier_Formation_ONG_Revise.docx");
}).catch(err => {
  console.error("Error generating document:", err);
});
