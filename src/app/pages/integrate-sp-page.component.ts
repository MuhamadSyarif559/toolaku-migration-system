import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';

interface SheetData {
  name: string;
  headers: string[];
  rows: Record<string, unknown>[];
}

interface WorkbookData {
  fileName: string;
  sheets: SheetData[];
}

interface FieldBinding {
  enabled: boolean;
  targetField: string;
  sourceField: string;
  defaultValue: string;
}

@Component({
  selector: 'app-integrate-sp-page',
  imports: [CommonModule, FormsModule],
  templateUrl: './integrate-sp-page.component.html',
  styleUrl: './integrate-sp-page.component.css'
})
export class IntegrateSpPageComponent {
  spMode: 'stored_procedure' | 'sharepoint' = 'stored_procedure';
  operationName = '';

  workbook: WorkbookData | null = null;
  selectedSheet = '';
  startRow = 2;

  storedProcedureFileName = '';
  storedProcedureText = '';
  procedureName = '';
  procedureParameters: string[] = [];

  fieldBindings: FieldBinding[] = [];

  busy = false;
  infoMessage = '';
  warningMessage = '';
  executionSummary = '';
  generatedSql = '';

  async onStoredProcedureFileSelected(event: Event): Promise<void> {
    const file = this.extractFile(event);
    if (!file) {
      return;
    }

    this.infoMessage = '';
    this.warningMessage = '';
    this.executionSummary = '';
    this.generatedSql = '';

    try {
      this.busy = true;
      this.storedProcedureText = await file.text();
      this.storedProcedureFileName = file.name;
      this.analyzeStoredProcedure();
      this.rebuildFieldBindings();
      if (this.procedureName) {
        this.infoMessage = `Loaded stored procedure file: ${file.name}`;
      } else {
        this.warningMessage = 'File uploaded, but no CREATE/ALTER PROCEDURE statement was detected.';
      }
    } catch (error) {
      this.warningMessage = this.toErrorMessage(error, 'Could not read stored procedure file.');
      this.storedProcedureText = '';
      this.storedProcedureFileName = '';
      this.procedureName = '';
      this.procedureParameters = [];
      this.fieldBindings = [];
    } finally {
      this.busy = false;
    }
  }

  async onExcelFileSelected(event: Event): Promise<void> {
    const file = this.extractFile(event);
    if (!file) {
      return;
    }

    this.infoMessage = '';
    this.warningMessage = '';
    this.executionSummary = '';
    this.generatedSql = '';

    try {
      this.busy = true;
      this.workbook = await this.parseWorkbook(file);
      this.selectedSheet = this.workbook.sheets[0]?.name ?? '';
      this.startRow = 2;
      this.rebuildFieldBindings();
      this.infoMessage = `Loaded Excel file: ${file.name}`;
    } catch (error) {
      this.warningMessage = this.toErrorMessage(error, 'Could not read Excel file.');
      this.workbook = null;
      this.selectedSheet = '';
      this.fieldBindings = [];
    } finally {
      this.busy = false;
    }
  }

  onSheetChange(): void {
    this.startRow = 2;
    this.generatedSql = '';
    this.rebuildFieldBindings();
  }

  generateSqlScript(): void {
    this.infoMessage = '';
    this.warningMessage = '';
    this.executionSummary = '';
    this.generatedSql = '';

    if (this.spMode !== 'stored_procedure') {
      this.warningMessage = 'SQL generation is only available for Stored Procedure mode.';
      return;
    }

    if (!this.procedureName.trim()) {
      this.warningMessage = 'Please upload a stored procedure SQL file first.';
      return;
    }

    if (!this.selectedSheetData) {
      this.warningMessage = 'Please upload Excel and choose a sheet first.';
      return;
    }

    if (this.startRow < 2) {
      this.warningMessage = 'Start row must be at least 2 (row 1 is header).';
      return;
    }

    const rows = this.getRowsForExecution();
    if (!rows.length) {
      this.warningMessage = 'No data rows found from selected start row.';
      return;
    }

    const activeBindings = this.fieldBindings.filter((binding) => binding.enabled && binding.targetField.trim());
    if (!activeBindings.length) {
      this.warningMessage = 'No active parameter mappings. Please map at least one SP parameter.';
      return;
    }

    this.generatedSql = this.buildLoopSqlScript(rows, activeBindings);
    this.executionSummary = `Generated loop SQL for ${rows.length} row(s).`;
    this.infoMessage = 'Loop SQL script generated successfully. Copy and run it in SQL Server.';
  }

  get selectedSheetData(): SheetData | undefined {
    return this.workbook?.sheets.find((sheet) => sheet.name === this.selectedSheet);
  }

  get excelHeaders(): string[] {
    return this.selectedSheetData?.headers ?? [];
  }

  get previewRows(): Record<string, unknown>[] {
    const activeBindings = this.fieldBindings.filter((binding) => binding.enabled && binding.targetField.trim());
    if (!activeBindings.length) {
      return [];
    }

    return this.getRowsForExecution()
      .slice(0, 3)
      .map((row) => {
        const mapped: Record<string, unknown> = {};
        activeBindings.forEach((binding) => {
          const sourceValue = binding.sourceField.trim() ? row[binding.sourceField] : '';
          mapped[this.toParameterName(binding.targetField)] = this.isEmptyValue(sourceValue)
            ? binding.defaultValue
            : sourceValue;
        });
        return mapped;
      });
  }

  get queryPreview(): string {
    const activeBindings = this.fieldBindings.filter((binding) => binding.enabled && binding.targetField.trim());
    if (!activeBindings.length || !this.procedureName.trim()) {
      return 'Upload SP + Excel and set parameter mapping to preview SQL.';
    }

    const sampleRow = this.getRowsForExecution()[0];
    if (!sampleRow) {
      return `EXEC ${this.procedureName}\n  -- no data rows from selected start row`;
    }

    const assignments = activeBindings.map((binding) => {
      const sourceValue = binding.sourceField.trim() ? sampleRow[binding.sourceField] : '';
      const resolvedValue = this.isEmptyValue(sourceValue) ? binding.defaultValue : sourceValue;
      return `${this.toParameterName(binding.targetField)} = ${this.toSqlLiteral(resolvedValue)}`;
    });

    const name = this.operationName.trim() ? ` -- ${this.operationName.trim()}` : '';
    return [`EXEC ${this.procedureName}${name}`, `  ${assignments.join(',\n  ')};`].join('\n');
  }

  private analyzeStoredProcedure(): void {
    const sql = this.storedProcedureText;
    this.procedureName = this.extractProcedureName(sql);
    this.procedureParameters = this.extractParameterNames(sql);
  }

  private extractProcedureName(sql: string): string {
    const cleaned = sql.replace(/--.*$/gm, '');
    const definitionRegex = /\b(?:create(?:\s+or\s+alter)?|alter)\s+proc(?:edure)?\s+((?:\[[^\]]+\]|[A-Za-z_][\w$#@]*)\s*(?:\.\s*(?:\[[^\]]+\]|[A-Za-z_][\w$#@]*))*)/i;
    const definitionMatch = cleaned.match(definitionRegex);
    if (definitionMatch?.[1]) {
      return definitionMatch[1].replace(/\s+/g, '').trim();
    }

    const execMatch = cleaned.match(/\bexec(?:ute)?\s+((?:\[[^\]]+\]|[A-Za-z_][\w$#@]*)\s*(?:\.\s*(?:\[[^\]]+\]|[A-Za-z_][\w$#@]*))*)/i);
    return execMatch?.[1]?.replace(/\s+/g, '').trim() ?? '';
  }

  private extractParameterNames(sql: string): string[] {
    const names: string[] = [];
    const seen = new Set<string>();

    const cleaned = sql.replace(/--.*$/gm, '');
    const definitionPartMatch = cleaned.match(/\b(?:create(?:\s+or\s+alter)?|alter)\s+proc(?:edure)?[\s\S]*?\bAS\b/i);
    const parameterSource = definitionPartMatch ? definitionPartMatch[0] : cleaned;

    const regex = /@([A-Za-z_][A-Za-z0-9_]*)/g;
    let match = regex.exec(parameterSource);
    while (match) {
      const normalized = `@${match[1]}`;
      const key = normalized.toLowerCase();
      if (!seen.has(key)) {
        seen.add(key);
        names.push(normalized);
      }
      match = regex.exec(parameterSource);
    }

    return names;
  }

  private rebuildFieldBindings(): void {
    const headers = this.excelHeaders;

    if (this.procedureParameters.length) {
      const previousByTarget = new Map(this.fieldBindings.map((binding) => [binding.targetField.toLowerCase(), binding]));

      this.fieldBindings = this.procedureParameters.map((parameter) => {
        const previous = previousByTarget.get(parameter.toLowerCase());
        const autoSource = this.findBestHeaderMatch(parameter, headers);

        return {
          enabled: previous?.enabled ?? true,
          targetField: parameter,
          sourceField: previous?.sourceField ?? autoSource,
          defaultValue: previous?.defaultValue ?? ''
        };
      });

      return;
    }

    const previousBySource = new Map(this.fieldBindings.map((binding) => [binding.sourceField, binding]));
    this.fieldBindings = headers.map((header) => {
      const previous = previousBySource.get(header);
      return {
        enabled: previous?.enabled ?? true,
        targetField: previous?.targetField || `@${header}`,
        sourceField: header,
        defaultValue: previous?.defaultValue ?? ''
      };
    });
  }

  private findBestHeaderMatch(parameter: string, headers: string[]): string {
    const normalizedParam = this.normalizeName(parameter.replace(/^@/, ''));
    return headers.find((header) => this.normalizeName(header) === normalizedParam) ?? '';
  }

  private normalizeName(value: string): string {
    return value.replace(/[^A-Za-z0-9]/g, '').toLowerCase();
  }

  private toParameterName(value: string): string {
    const trimmed = value.trim();
    if (!trimmed) {
      return '';
    }

    return trimmed.startsWith('@') ? trimmed : `@${trimmed}`;
  }

  private toColumnAlias(value: string): string {
    const raw = this.toParameterName(value).replace(/^@/, '');
    return raw.replace(/[^A-Za-z0-9_]/g, '_') || 'Param';
  }

  private buildLoopSqlScript(rows: Record<string, unknown>[], activeBindings: FieldBinding[]): string {
    const aliasCount = new Map<string, number>();
    const columns = activeBindings.map((binding) => {
      const parameter = this.toParameterName(binding.targetField);
      const baseAlias = this.toColumnAlias(binding.targetField);
      const count = aliasCount.get(baseAlias) ?? 0;
      aliasCount.set(baseAlias, count + 1);
      const alias = count === 0 ? baseAlias : `${baseAlias}_${count + 1}`;

      return {
        parameter,
        alias,
        sourceField: binding.sourceField,
        defaultValue: binding.defaultValue
      };
    });

    const declareTableColumns = columns.map((column) => `  [${column.alias}] NVARCHAR(MAX) NULL`).join(',\n');

    const insertColumnList = columns.map((column) => `[${column.alias}]`).join(', ');
    const insertRows = rows
      .map((row) => {
        const values = columns.map((column) => {
          const sourceValue = column.sourceField.trim() ? row[column.sourceField] : '';
          const resolvedValue = this.isEmptyValue(sourceValue) ? column.defaultValue : sourceValue;
          return this.toSqlLiteral(resolvedValue);
        });

        return `(${values.join(', ')})`;
      })
      .join(',\n');

    const declareExecVariables = columns
      .map((column) => `DECLARE ${column.parameter} NVARCHAR(MAX);`)
      .join('\n');

    const setExecVariables = columns
      .map((column) => `  SET ${column.parameter} = (SELECT [${column.alias}] FROM ExcelData WHERE RowId = @CurrentRowId);`)
      .join('\n');

    const execAssignments = columns
      .map((column) => `    ${column.parameter} = ${column.parameter}`)
      .join(',\n');

    return [
      '-- Generated SQL loop from Excel upload',
      'SET NOCOUNT ON;',
      '',
      'IF OBJECT_ID(\'ExcelData\') IS NOT NULL DROP TABLE ExcelData;',
      '',
      'CREATE TABLE ExcelData (',
      '  RowId INT IDENTITY(1,1) NOT NULL PRIMARY KEY,',
      `${declareTableColumns}`,
      ');',
      '',
      `INSERT INTO ExcelData (${insertColumnList})`,
      'VALUES',
      `${insertRows};`,
      '',
      'DECLARE @CurrentRowId INT = 1;',
      'DECLARE @MaxRowId INT;',
      'SELECT @MaxRowId = MAX(RowId) FROM ExcelData;',
      '',
      `${declareExecVariables}`,
      '',
      'WHILE @CurrentRowId <= ISNULL(@MaxRowId, 0)',
      'BEGIN',
      `${setExecVariables}`,
      '',
      `  EXEC ${this.procedureName}`,
      `${execAssignments};`,
      '',
      '  SET @CurrentRowId = @CurrentRowId + 1;',
      'END;',
      '',
      'DROP TABLE ExcelData;'
    ].join('\n');
  }

  private toSqlLiteral(value: unknown): string {
    if (this.isEmptyValue(value)) {
      return 'NULL';
    }

    if (typeof value === 'number' && Number.isFinite(value)) {
      return String(value);
    }

    if (typeof value === 'boolean') {
      return value ? '1' : '0';
    }

    const text = String(value).trim();
    if (/^-?\d+(\.\d+)?$/.test(text)) {
      return text;
    }

    const escaped = text.replace(/'/g, "''");
    return `N'${escaped}'`;
  }

  private getRowsForExecution(): Record<string, unknown>[] {
    const sheet = this.selectedSheetData;
    if (!sheet) {
      return [];
    }

    const startIndex = Math.max(this.startRow - 2, 0);
    return sheet.rows.slice(startIndex);
  }

  private extractFile(event: Event): File | null {
    const input = event.target as HTMLInputElement;
    return input.files?.item(0) ?? null;
  }

  private async parseWorkbook(file: File): Promise<WorkbookData> {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

    const sheets = workbook.SheetNames.map((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json<(string | number | boolean | null)[]>(worksheet, {
        header: 1,
        defval: ''
      });

      return this.toSheetData(sheetName, matrix);
    });

    return {
      fileName: file.name,
      sheets
    };
  }

  // AFTER
    private toSheetData(name: string, matrix: (string | number | boolean | null | Date)[][]): SheetData {
    const headerRowIndex = matrix.findIndex((row) => row.some((cell) => String(cell ?? '').trim() !== ''));
    if (headerRowIndex < 0) {
      return { name, headers: [], rows: [] };
    }

    const rawHeaders = matrix[headerRowIndex] ?? [];
    const headers = this.normalizeHeaders(rawHeaders);

    const rows = matrix
      .slice(headerRowIndex + 1)
      .filter((row) => row.some((cell) => String(cell ?? '').trim() !== ''))
      // AFTER
      .map((row) => {
        const record: Record<string, unknown> = {};
        headers.forEach((header, index) => {
          const cell = row[index] ?? '';
          if (header === 'Date' || header === 'DueDate' || header === 'RequireDate' || header === 'PostDate') {
            console.log(`DATE FIELD >> Header: ${header} | Value: "${cell}" | Type: ${typeof cell}`);
          }
          record[header] = this.parseCellValue(cell);
        });
        return record;
      });

    return { name, headers, rows };
  }

    private parseCellValue(cell: string | number | boolean | null | Date): unknown {
      if (cell instanceof Date) {
        return this.toIsoDateString(cell);
      }
    
      // cellDates: true sometimes passes Date objects that lost their prototype
      if (typeof cell === 'object' && cell !== null) {
        const d = new Date(cell as unknown as string);
        if (!isNaN(d.getTime())) {
          return this.toIsoDateString(d);
        }
      }
    
      if (typeof cell === 'number' && cell > 25569 && cell < 60000) {
        const date = new Date(Math.round((cell - 25569) * 86400 * 1000));
        return this.toIsoDateString(date);
      }
    
      return cell;
    }

  // AFTER
  private normalizeHeaders(values: (string | number | boolean | null | Date)[]): string[] {
    const seen = new Map<string, number>();

    return values.map((value, index) => {
      const base = String(value ?? '').trim() || `Column ${index + 1}`;
      const count = seen.get(base) ?? 0;
      seen.set(base, count + 1);

      return count === 0 ? base : `${base}_${count + 1}`;
    });
  }

  private isEmptyValue(value: unknown): boolean {
    return value === null || value === undefined || String(value).trim() === '';
  }

  // AFTER
  private toIsoDateString(date: Date): string {
    const adjusted = new Date(date.getTime() + 24 * 60 * 60 * 1000);
    const yyyy = adjusted.getUTCFullYear();
    const mm = String(adjusted.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(adjusted.getUTCDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  private toErrorMessage(error: unknown, fallback: string): string {
    if (error instanceof Error && error.message) {
      return error.message;
    }

    return fallback;
  }
}
