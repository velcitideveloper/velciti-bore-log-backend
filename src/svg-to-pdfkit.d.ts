declare module 'svg-to-pdfkit' {
  import type PDFDocument from 'pdfkit';

  interface SVGtoPDFOptions {
    width?: number;
    height?: number;
    preserveAspectRatio?: string;
    useCSS?: boolean;
    fontCallback?: (
      family: string,
      bold: boolean,
      italic: boolean,
      fontOptions: Record<string, unknown>,
    ) => string;
    imageCallback?: (uri: string) => Buffer | string;
    documentCallback?: (
      uri: string,
      callback: (err: Error | null, data?: unknown) => void,
    ) => void;
    colorCallback?: (
      type: string,
      color: string,
      alpha: number,
    ) => [string, number] | null;
  }

  function SVGtoPDF(
    doc: InstanceType<typeof PDFDocument>,
    svg: string,
    x: number,
    y: number,
    options?: SVGtoPDFOptions,
  ): void;

  export = SVGtoPDF;
}
