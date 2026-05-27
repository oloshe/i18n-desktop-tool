import {
  parseExcelBuffer,
  readExcelWorkbookInfo,
  rowsFromExcelBuffer,
  type ExcelReadOptions
} from "../core/excelParser";

type ExcelWorkerRequest =
  | { id: number; type: "workbookInfo"; buffer: ArrayBuffer; sheetName?: string }
  | { id: number; type: "preview"; buffer: ArrayBuffer; options: ExcelReadOptions }
  | { id: number; type: "rows"; buffer: ArrayBuffer; options: ExcelReadOptions };

type ExcelWorkerResponse =
  | { id: number; ok: true; payload: unknown }
  | { id: number; ok: false; error: string };

self.onmessage = (event: MessageEvent<ExcelWorkerRequest>) => {
  const request = event.data;

  try {
    const payload =
      request.type === "workbookInfo"
        ? readExcelWorkbookInfo(request.buffer, request.sheetName)
        : request.type === "preview"
          ? parseExcelBuffer(request.buffer, request.options)
          : rowsFromExcelBuffer(request.buffer, request.options);

    self.postMessage({ id: request.id, ok: true, payload } satisfies ExcelWorkerResponse);
  } catch (error) {
    self.postMessage({
      id: request.id,
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    } satisfies ExcelWorkerResponse);
  }
};
