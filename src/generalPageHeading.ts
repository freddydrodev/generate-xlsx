import { IXlsxHeadingTextContent } from "./types";

export const generalPageHeading: (
  headerLength: number,
  bankName: string
) => IXlsxHeadingTextContent[][] = (headerLength, bankName) => [
  [
    {
      text: "MINISTERE DES TRANSPORTS",
      startAt: 1,
      endAt: 3,
    },
    {
      text: "REPUBLIQUE DE COTE D'IVOIRE",
      startAt: headerLength - 1,
      endAt: headerLength,
      centered: true,
    },
  ],
  [
    {
      text: "______________",
      startAt: 1,
      endAt: 3,
    },
    {
      text: "Union - Discipline - Travail",
      startAt: headerLength - 1,
      endAt: headerLength,
      centered: true,
    },
  ],
  [
    {
      text: "MINISTERE DELEGUE AUPRES DU MINISTRE	",
      startAt: 1,
      endAt: 3,
    },
    {
      text: "_____________",
      startAt: headerLength - 1,
      endAt: headerLength,
      centered: true,
    },
  ],
  [
    {
      text: "DES TRANSPORTS CHARGE DES AFFAIRES	",
      startAt: 1,
      endAt: 3,
    },
  ],
  [
    {
      text: "MARITIMES		",
      startAt: 1,
      endAt: 3,
    },
  ],
  [
    {
      text: "_____________",
      startAt: 1,
      endAt: 3,
    },
  ],
  [
    {
      text: "DIRECTION GENERALE DES AFFAIRES	",
      startAt: 1,
      endAt: 3,
    },
  ],
  [
    {
      text: "MARITIMES ET PORTUAIRES	",
      startAt: 1,
      endAt: 3,
    },
  ],
  [
    {
      text: "_____________",
      startAt: 1,
      endAt: 3,
    },
  ],
  [
    {
      text: "DIRECTION DE L'INTENDANCE",
      startAt: 1,
      endAt: 3,
    },
  ],
  [
    {
      text: "ET DE LA FACTURATION",
      startAt: 1,
      endAt: 3,
    },
    {
      text: "Abidjan le,",
      startAt: headerLength - 1,
      endAt: headerLength,
    },
  ],
  ...(Array(2) as any).fill([]),
  [
    {
      text: bankName,
      startAt: 3,
      endAt: headerLength - 1,
      centered: true,
      title: true,
    },
  ],
  ...(Array(2) as any).fill([]),
];
