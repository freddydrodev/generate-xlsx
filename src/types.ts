export interface IXlsxHeadingTextContent {
  text?: any;
  formula?: string;
  startAt: number;
  endAt: number;
  centered?: boolean;
  title?: "top" | "bottom" | true;
  subTitle?: boolean;
  pageBreak?: boolean;
  bordered?: boolean;
  bold?: boolean;
  bgColor?: string;
}
