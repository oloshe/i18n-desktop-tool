use calamine::{open_workbook_auto, open_workbook_auto_from_rs, Data, Reader, Sheets};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::io::Cursor;

const PREVIEW_LIMIT: usize = 20;

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct ExcelReadOptions {
  sheet_name: Option<String>,
  sheet_names: Option<Vec<String>>,
  skip_rows: Option<usize>,
  header_row: Option<usize>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ExcelWorkbookInfo {
  sheet_names: Vec<String>,
  active_sheet_name: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ExcelPreview {
  sheet_names: Vec<String>,
  active_sheet_name: String,
  header_row: usize,
  skip_rows: usize,
  headers: Vec<String>,
  rows: Vec<HashMap<String, String>>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct ExcelHealthCheckRowRef {
  sheet_name: String,
  row_number: usize,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ExcelDuplicateKeyIssue {
  key: String,
  locations: Vec<ExcelHealthCheckRowRef>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ExcelMissingTranslationIssue {
  sheet_name: String,
  row_number: usize,
  key: String,
  lang: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct ExcelHealthCheckResult {
  duplicate_keys: Vec<ExcelDuplicateKeyIssue>,
  empty_key_rows: Vec<ExcelHealthCheckRowRef>,
  missing_translations: Vec<ExcelMissingTranslationIssue>,
}

#[tauri::command]
async fn read_excel_workbook_info(path: String, sheet_name: Option<String>) -> Result<ExcelWorkbookInfo, String> {
  tauri::async_runtime::spawn_blocking(move || {
    let mut workbook = open_workbook_auto(path).map_err(|error| error.to_string())?;
    workbook_info(&mut workbook, sheet_name.as_deref())
  })
  .await
  .map_err(|error| error.to_string())?
}

#[tauri::command]
async fn read_excel_workbook_info_bytes(bytes: Vec<u8>, sheet_name: Option<String>) -> Result<ExcelWorkbookInfo, String> {
  tauri::async_runtime::spawn_blocking(move || {
    let cursor = Cursor::new(bytes);
    let mut workbook = open_workbook_auto_from_rs(cursor).map_err(|error| error.to_string())?;
    workbook_info(&mut workbook, sheet_name.as_deref())
  })
  .await
  .map_err(|error| error.to_string())?
}

#[tauri::command]
async fn preview_excel(path: String, options: ExcelReadOptions) -> Result<ExcelPreview, String> {
  tauri::async_runtime::spawn_blocking(move || {
    let mut workbook = open_workbook_auto(path).map_err(|error| error.to_string())?;
    parse_preview(&mut workbook, options)
  })
  .await
  .map_err(|error| error.to_string())?
}

#[tauri::command]
async fn preview_excel_bytes(bytes: Vec<u8>, options: ExcelReadOptions) -> Result<ExcelPreview, String> {
  tauri::async_runtime::spawn_blocking(move || {
    let cursor = Cursor::new(bytes);
    let mut workbook = open_workbook_auto_from_rs(cursor).map_err(|error| error.to_string())?;
    parse_preview(&mut workbook, options)
  })
  .await
  .map_err(|error| error.to_string())?
}

#[tauri::command]
async fn rows_from_excel(path: String, options: ExcelReadOptions) -> Result<Vec<HashMap<String, String>>, String> {
  tauri::async_runtime::spawn_blocking(move || {
    let mut workbook = open_workbook_auto(path).map_err(|error| error.to_string())?;
    parse_rows(&mut workbook, options)
  })
  .await
  .map_err(|error| error.to_string())?
}

#[tauri::command]
async fn rows_from_excel_bytes(bytes: Vec<u8>, options: ExcelReadOptions) -> Result<Vec<HashMap<String, String>>, String> {
  tauri::async_runtime::spawn_blocking(move || {
    let cursor = Cursor::new(bytes);
    let mut workbook = open_workbook_auto_from_rs(cursor).map_err(|error| error.to_string())?;
    parse_rows(&mut workbook, options)
  })
  .await
  .map_err(|error| error.to_string())?
}

#[tauri::command]
async fn inspect_excel(
  path: String,
  key_column: String,
  language_columns: HashMap<String, String>,
  options: ExcelReadOptions,
) -> Result<ExcelHealthCheckResult, String> {
  tauri::async_runtime::spawn_blocking(move || {
    let mut workbook = open_workbook_auto(path).map_err(|error| error.to_string())?;
    inspect_workbook(&mut workbook, &key_column, &language_columns, options)
  })
  .await
  .map_err(|error| error.to_string())?
}

fn workbook_info<RS>(workbook: &mut Sheets<RS>, sheet_name: Option<&str>) -> Result<ExcelWorkbookInfo, String>
where
  RS: std::io::Read + std::io::Seek,
{
  let sheet_names = workbook.sheet_names().to_owned();
  let active_sheet_name = active_sheet_name(&sheet_names, sheet_name)?;
  Ok(ExcelWorkbookInfo { sheet_names, active_sheet_name })
}

fn parse_preview<RS>(workbook: &mut Sheets<RS>, options: ExcelReadOptions) -> Result<ExcelPreview, String>
where
  RS: std::io::Read + std::io::Seek,
{
  let skip_rows = options.skip_rows.unwrap_or(0);
  let header_row = options.header_row.unwrap_or(1).max(1);
  let sheet_names = workbook.sheet_names().to_owned();
  let active_sheet_name = active_sheet_name(&sheet_names, options.sheet_name.as_deref())?;
  let matrix = sheet_matrix(workbook, &active_sheet_name)?;
  let headers = get_headers(&matrix, header_row)?;
  let data_start_index = skip_rows.max(header_row);

  let rows = matrix
    .iter()
    .skip(data_start_index)
    .filter(|row| is_importable_row(row, headers.len()))
    .take(PREVIEW_LIMIT)
    .map(|row| row_to_record(row, &headers))
    .collect();

  Ok(ExcelPreview { sheet_names, active_sheet_name, header_row, skip_rows, headers, rows })
}

fn parse_rows<RS>(workbook: &mut Sheets<RS>, options: ExcelReadOptions) -> Result<Vec<HashMap<String, String>>, String>
where
  RS: std::io::Read + std::io::Seek,
{
  let skip_rows = options.skip_rows.unwrap_or(0);
  let header_row = options.header_row.unwrap_or(1).max(1);
  let sheet_names = workbook.sheet_names().to_owned();
  let data_start_index = skip_rows.max(header_row);
  let mut rows = Vec::new();

  for sheet_name in active_sheet_names(&sheet_names, options.sheet_name.as_deref(), options.sheet_names.as_deref())? {
    let matrix = sheet_matrix(workbook, &sheet_name)?;
    let headers = get_headers(&matrix, header_row)?;
    rows.extend(
      matrix
        .iter()
        .skip(data_start_index)
        .filter(|row| has_filled_cell(row, headers.len()))
        .map(|row| row_to_record(row, &headers)),
    );
  }

  Ok(rows)
}

fn inspect_workbook<RS>(
  workbook: &mut Sheets<RS>,
  key_column: &str,
  language_columns: &HashMap<String, String>,
  options: ExcelReadOptions,
) -> Result<ExcelHealthCheckResult, String>
where
  RS: std::io::Read + std::io::Seek,
{
  if key_column.trim().is_empty() {
    return Err("Key column is required.".to_string());
  }

  let languages: Vec<(&String, &String)> = language_columns
    .iter()
    .filter(|(lang, column)| !lang.trim().is_empty() && !column.trim().is_empty())
    .collect();
  if languages.is_empty() {
    return Err("At least one language column is required.".to_string());
  }

  let skip_rows = options.skip_rows.unwrap_or(0);
  let header_row = options.header_row.unwrap_or(1).max(1);
  let sheet_names = workbook.sheet_names().to_owned();
  let data_start_index = skip_rows.max(header_row);
  let mut duplicate_map: HashMap<String, Vec<ExcelHealthCheckRowRef>> = HashMap::new();
  let mut empty_key_rows = Vec::new();
  let mut missing_translations = Vec::new();

  for sheet_name in active_sheet_names(&sheet_names, options.sheet_name.as_deref(), options.sheet_names.as_deref())? {
    let matrix = sheet_matrix(workbook, &sheet_name)?;
    let headers = get_headers(&matrix, header_row)?;
    let key_index = column_index(&headers, key_column, &sheet_name)?;
    let language_indexes: Vec<(&String, usize)> = languages
      .iter()
      .map(|(lang, column)| column_index(&headers, column, &sheet_name).map(|index| (*lang, index)))
      .collect::<Result<Vec<_>, _>>()?;

    for (offset, row) in matrix.iter().skip(data_start_index).enumerate() {
      let row_number = data_start_index + offset + 1;
      if !has_filled_cell(row, headers.len()) {
        continue;
      }

      let key = row.get(key_index).map(|value| value.trim()).unwrap_or_default();
      if key.is_empty() {
        empty_key_rows.push(ExcelHealthCheckRowRef { sheet_name: sheet_name.clone(), row_number });
        continue;
      }

      let filled_language_count = language_indexes
        .iter()
        .filter(|(_, index)| row.get(*index).map(|value| !value.trim().is_empty()).unwrap_or(false))
        .count();
      if filled_language_count == 0 {
        continue;
      }

      duplicate_map
        .entry(key.to_string())
        .or_default()
        .push(ExcelHealthCheckRowRef { sheet_name: sheet_name.clone(), row_number });

      for (lang, index) in &language_indexes {
        if row.get(*index).map(|value| value.trim().is_empty()).unwrap_or(true) {
          missing_translations.push(ExcelMissingTranslationIssue {
            sheet_name: sheet_name.clone(),
            row_number,
            key: key.to_string(),
            lang: (*lang).to_string(),
          });
        }
      }
    }
  }

  let duplicate_keys = duplicate_map
    .into_iter()
    .filter(|(_, locations)| locations.len() > 1)
    .map(|(key, locations)| ExcelDuplicateKeyIssue { key, locations })
    .collect();

  Ok(ExcelHealthCheckResult { duplicate_keys, empty_key_rows, missing_translations })
}

fn active_sheet_name(sheet_names: &[String], requested: Option<&str>) -> Result<String, String> {
  if sheet_names.is_empty() {
    return Err("Excel file has no readable sheets.".to_string());
  }

  Ok(
    requested
      .filter(|name| sheet_names.iter().any(|sheet| sheet == name))
      .unwrap_or(&sheet_names[0])
      .to_string(),
  )
}

fn active_sheet_names(sheet_names: &[String], requested: Option<&str>, requested_names: Option<&[String]>) -> Result<Vec<String>, String> {
  if sheet_names.is_empty() {
    return Err("Excel file has no readable sheets.".to_string());
  }

  let mut selected = Vec::new();
  if let Some(names) = requested_names {
    for name in names {
      if sheet_names.iter().any(|sheet| sheet == name) && !selected.iter().any(|sheet| sheet == name) {
        selected.push(name.to_string());
      }
    }
  }

  if !selected.is_empty() {
    return Ok(selected);
  }

  Ok(vec![active_sheet_name(sheet_names, requested)?])
}

fn sheet_matrix<RS>(workbook: &mut Sheets<RS>, sheet_name: &str) -> Result<Vec<Vec<String>>, String>
where
  RS: std::io::Read + std::io::Seek,
{
  let range = workbook.worksheet_range(sheet_name).map_err(|error| error.to_string())?;
  let matrix: Vec<Vec<String>> = range
    .rows()
    .map(|row| row.iter().map(normalize_cell).collect())
    .collect();

  if matrix.is_empty() {
    return Err(format!("Sheet \"{}\" is empty.", sheet_name));
  }

  Ok(matrix)
}

fn get_headers(matrix: &[Vec<String>], header_row: usize) -> Result<Vec<String>, String> {
  let header_index = header_row - 1;
  let row = matrix
    .get(header_index)
    .ok_or_else(|| "Header row is outside the Excel content range.".to_string())?;
  let headers: Vec<String> = row.iter().map(|value| value.trim().to_string()).filter(|value| !value.is_empty()).collect();

  if headers.is_empty() {
    return Err("Header row has no valid headers.".to_string());
  }

  for (index, header) in headers.iter().enumerate() {
    if headers.iter().skip(index + 1).any(|other| other == header) {
      return Err("Excel headers contain duplicate column names.".to_string());
    }
  }

  Ok(headers)
}

fn row_to_record(row: &[String], headers: &[String]) -> HashMap<String, String> {
  headers
    .iter()
    .enumerate()
    .map(|(index, header)| (header.to_string(), row.get(index).cloned().unwrap_or_default()))
    .collect()
}

fn column_index(headers: &[String], column: &str, sheet_name: &str) -> Result<usize, String> {
  headers
    .iter()
    .position(|header| header == column)
    .ok_or_else(|| format!("Sheet \"{}\" is missing column \"{}\".", sheet_name, column))
}

fn is_importable_row(row: &[String], width: usize) -> bool {
  row.iter().take(width).filter(|value| !value.trim().is_empty()).count() > 1
}

fn has_filled_cell(row: &[String], width: usize) -> bool {
  row.iter().take(width).any(|value| !value.trim().is_empty())
}

fn normalize_cell(cell: &Data) -> String {
  match cell {
    Data::Empty => String::new(),
    Data::Float(value) if value.fract() == 0.0 => format!("{value:.0}"),
    Data::DateTime(value) => value.to_string(),
    Data::DateTimeIso(value) | Data::DurationIso(value) | Data::String(value) => value.trim().to_string(),
    _ => cell.to_string().trim().to_string(),
  }
}

pub fn run() {
  tauri::Builder::default()
    .plugin(tauri_plugin_dialog::init())
    .plugin(tauri_plugin_fs::init())
    .invoke_handler(tauri::generate_handler![
      read_excel_workbook_info,
      read_excel_workbook_info_bytes,
      preview_excel,
      preview_excel_bytes,
      rows_from_excel,
      rows_from_excel_bytes,
      inspect_excel
    ])
    .run(tauri::generate_context!())
    .expect("error while running tauri application");
}
