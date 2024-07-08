import ExcelJS, { Alignment, Font } from 'exceljs';

type GenerateXLSXArgs = {
    fileName: string;
    config: {
        name: string;
        orientation?: "portrait" | "landscape";
        zoom?: number;
        colWidth?: number;
        colHeight?: number;
    };
    data: {
        [key: string]: string | number;
    }[];
    headers: Pick<Partial<ExcelJS.Column>, "key" | "header" | "border" | "alignment" | "fill" | "numFmt" | "values" | "width" | "style" | "font">[];
    rowAlignment?: Partial<Alignment>;
    defaultFont?: Partial<Font>;
    boldFont?: Partial<Font>;
    height?: number;
    title?: {
        height?: number;
        fontSize?: number;
    };
};
declare const generateXLSXGrid: (args: GenerateXLSXArgs) => Promise<void>;

export { generateXLSXGrid };
