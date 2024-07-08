import type { Font, Borders, Alignment } from "exceljs";

export const DEFAULT_BORDER: Partial<Borders> = {
  bottom: { color: { argb: "FF000000" }, style: "thin" },
  top: { color: { argb: "FF000000" }, style: "thin" },
  left: { color: { argb: "FF000000" }, style: "thin" },
  right: { color: { argb: "FF000000" }, style: "thin" },
};

export const DEFAULT_ROW_ALIGNEMENT: Partial<Alignment> = {
  horizontal: "center",
  vertical: "middle",
  wrapText: true,
};

export const DEFAULT_XLSX_FONT = "Goudy Old Style";

export const BOLD_XLSX_FONT = "Goudy Old Style Bold";

export const ROW_HEIGHT = 50;

export const XLS_JUSTIF_HEADING_HEIGHT = 70;

export const DEFAULT_ROW_FONT: Partial<Font> = {
  size: 16,
  name: DEFAULT_XLSX_FONT,
};

export const BOLD_FONT: Partial<Font> = {
  ...DEFAULT_ROW_FONT,
  bold: true,
  name: BOLD_XLSX_FONT,
};

export const DEFAULT_HEADER_ROW_HEIGHT = 70;

export const DEFAULT_ROW_HEIGHT = 65;

export const DEFAULT_HEADLINE_HEIGHT = 25;

export const TABLE_LINE_PER_PAGE = 15;

export const TABLE_LINE_ON_FIRST_PAGE = 9;

export const TABLE_HEADERS: any = {
  nom_complet: "NOM ET PRENOMS",
  matricule: "MLE",
  emploi: "EMPLOI",
  fonction: "FONCTION",
  annee_de_prise_de_fonction: "PRISE DE FONCTION AFF/MAR",
  anciennete: "ANCIENNETE",
  grade: "CATEGORIE",
  points: "POINTS",
  prime_de_base: "PRIME BASE",
  prime_de_responsabilite: "PRIME RESPONSABILITE",
  prime_anciennete: "PRIME ANCIENNETE",
  prime_specifique: "PRIME SPECIFIQUE AFF MAR",
  relicat: "RETENUE",
  forfait: "FORFAIT",
  total: "TOTAL",
  taux: "TAUX DE LA PRIME DE BASE",
  prime_lie_a_l_emploie: "PRIME LIE A L'EMPLOI",
  prime_d_installation: "PRIME D'INSTALLATION",
  prime_de_logement: "PRIME DE LOGEMENT",
  prime_de_plus_value: "PRIME DE PLUS VALUE",
  prime_de_securite: "PRIME DE SECURITE",
  prime_de_travail_supplementaire: "PRIME DE TRAVAIL SUPPLEMENTAIRE",
  // rib: ,
};

export const getPeriod = (trimester: { nom: string; annee: number }) => {
  switch (trimester.nom) {
    case "Janvier - Mars":
      return `Du 01/01/${trimester.annee} au 31/03/${trimester.annee}`;
    case "Avril - Juin":
      return `Du 01/04/${trimester.annee} au 30/06/${trimester.annee}`;
    case "Juillet - Septembre":
      return `Du 01/07/${trimester.annee} au 30/09/${trimester.annee}`;
    case "Octobre - Décembre":
    default:
      return `Du 01/10/${trimester.annee} au 31/12/${trimester.annee}`;
  }
};

export const getPeriodMonth = (trimester: { nom: string; annee: number }) => {
  switch (trimester.nom) {
    case "Janvier - Mars":
      return `MARS ${trimester.annee}`;
    case "Avril - Juin":
      return `JUIN ${trimester.annee}`;
    case "Juillet - Septembre":
      return `SEPTEMBRE ${trimester.annee}`;
    case "Octobre - Décembre":
    default:
      return `DECEMBRE ${trimester.annee}`;
  }
};
