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
  workbook: XLSX.WorkBook;
  sheets: SheetData[];
}

type FallbackMode = 'when_source_empty' | 'always' | 'when_source_equals';
type TransformRule = 'none' | 'trim' | 'uppercase' | 'lowercase' | 'number' | 'date_iso';
type TargetType = 'any' | 'string' | 'number' | 'boolean' | 'date';

interface ColumnMapping {
  targetColumn: string;
  sourceColumn: string;
  defaultValue: string;
  fallbackMode: FallbackMode;
  fallbackCompareValue: string;
  transformRule: TransformRule;
  targetType: TargetType;
  required: boolean;
}

interface ValidationIssue {
  sheet: string;
  row: number;
  column: string;
  reason: string;
}

interface ErrorRow {
  sheet: string;
  row: number;
  reason: string;
  sourceRow: Record<string, unknown>;
}

interface DiffPreview {
  sheet: string;
  row: number;
  changes: string[];
}

interface SheetMigrationResult {
  processed: number;
  migrated: number;
  failed: number;
  skipped: number;
  issues: ValidationIssue[];
  errorRows: ErrorRow[];
  diffs: DiffPreview[];
  fillRates: Record<string, number>;
}

interface MigrationSummary {
  processed: number;
  migrated: number;
  failed: number;
  skipped: number;
  fillRates: Record<string, number>;
}

interface AuditLogEntry {
  timestamp: string;
  sourceFile: string;
  targetFile: string;
  batchMode: boolean;
  processed: number;
  migrated: number;
  failed: number;
}

@Component({
  selector: 'app-migration-page',
  imports: [CommonModule, FormsModule],
  templateUrl: './migration-page.component.html',
  styleUrl: './migration-page.component.css'
})
export class MigrationPageComponent {
  sourceWorkbook: WorkbookData | null = null;
  targetWorkbook: WorkbookData | null = null;

  selectedSourceSheet = '';
  selectedTargetSheet = '';

  mappings: ColumnMapping[] = [];

  runBatchMode = false;
  useKeyMergeMode = false;
  sourceKeyColumn = '';
  targetKeyColumn = '';
  downloadBackup = true;

  busy = false;
  errorMessage = '';
  infoMessage = '';

  validationIssues: ValidationIssue[] = [];
  errorRows: ErrorRow[] = [];
  diffPreview: DiffPreview[] = [];
  summary: MigrationSummary | null = null;
  auditLog: AuditLogEntry[] = [];

  readonly fallbackModes: { value: FallbackMode; label: string }[] = [
    { value: 'when_source_empty', label: 'Use default only when source is empty' },
    { value: 'always', label: 'Always use default value' },
    { value: 'when_source_equals', label: 'Use default when source equals value' }
  ];

  readonly transformRules: { value: TransformRule; label: string }[] = [
    { value: 'none', label: 'None' },
    { value: 'trim', label: 'Trim' },
    { value: 'uppercase', label: 'Uppercase' },
    { value: 'lowercase', label: 'Lowercase' },
    { value: 'number', label: 'Parse Number' },
    { value: 'date_iso', label: 'Date YYYY-MM-DD' }
  ];

  readonly targetTypes: { value: TargetType; label: string }[] = [
    { value: 'any', label: 'Any' },
    { value: 'string', label: 'String' },
    { value: 'number', label: 'Number' },
    { value: 'boolean', label: 'Boolean' },
    { value: 'date', label: 'Date' }
  ];

  async onSourceFileSelected(event: Event): Promise<void> {
    const file = this.extractFile(event);
    if (!file) {
      return;
    }

    this.clearMessages();

    try {
      this.busy = true;
      this.sourceWorkbook = await this.parseWorkbook(file);
      this.selectedSourceSheet = this.sourceWorkbook.sheets[0]?.name ?? '';
      this.sourceKeyColumn = '';
      this.rebuildMappings();
      this.infoMessage = `Loaded source file: ${file.name}`;
    } catch (error) {
      this.errorMessage = this.toErrorMessage(error, 'Could not read source file.');
      this.sourceWorkbook = null;
      this.selectedSourceSheet = '';
      this.sourceKeyColumn = '';
      this.mappings = [];
    } finally {
      this.busy = false;
    }
  }

  async onTargetFileSelected(event: Event): Promise<void> {
    const file = this.extractFile(event);
    if (!file) {
      return;
    }

    this.clearMessages();

    try {
      this.busy = true;
      this.targetWorkbook = await this.parseWorkbook(file);
      this.selectedTargetSheet = this.targetWorkbook.sheets[0]?.name ?? '';
      this.targetKeyColumn = '';
      this.rebuildMappings();
      this.infoMessage = `Loaded target file: ${file.name}`;
    } catch (error) {
      this.errorMessage = this.toErrorMessage(error, 'Could not read target file.');
      this.targetWorkbook = null;
      this.selectedTargetSheet = '';
      this.targetKeyColumn = '';
      this.mappings = [];
    } finally {
      this.busy = false;
    }
  }

  onSourceSheetChange(): void {
    this.sourceKeyColumn = '';
    this.rebuildMappings();
  }

  onTargetSheetChange(): void {
    this.targetKeyColumn = '';
    this.rebuildMappings();
  }

  canMigrate(): boolean {
    return Boolean(
      this.sourceSheet &&
      this.targetSheet &&
      this.mappings.some(
        (mapping) => mapping.sourceColumn || mapping.defaultValue.trim() !== '' || mapping.required
      )
    );
  }

  migrate(): void {
    this.clearMessages();
    this.validationIssues = [];
    this.errorRows = [];
    this.diffPreview = [];
    this.summary = null;

    if (!this.sourceWorkbook || !this.targetWorkbook || !this.sourceSheet || !this.targetSheet) {
      this.errorMessage = 'Please upload both files and choose source/target sheets.';
      return;
    }

    const activeMappings = this.mappings.filter(
      (mapping) => mapping.sourceColumn || mapping.defaultValue.trim() !== '' || mapping.required
    );

    if (!activeMappings.length) {
      this.errorMessage = 'Map at least one target column to a source column, default, or required rule.';
      return;
    }

    const destinationWorkbook = XLSX.read(
      XLSX.write(this.targetWorkbook.workbook, { type: 'array', bookType: 'xlsx' }),
      { type: 'array' }
    );

    const pairs = this.getSheetPairs();
    if (!pairs.length) {
      this.errorMessage = 'No matching sheet pairs found for migration.';
      return;
    }

    const summary: MigrationSummary = {
      processed: 0,
      migrated: 0,
      failed: 0,
      skipped: 0,
      fillRates: {}
    };

    for (const pair of pairs) {
      const mappings = this.runBatchMode
        ? this.buildAutoMappings(pair.source.headers, pair.target.headers)
        : activeMappings;

      const result = this.migrateSheet(pair.source, pair.target, mappings, destinationWorkbook);

      summary.processed += result.processed;
      summary.migrated += result.migrated;
      summary.failed += result.failed;
      summary.skipped += result.skipped;
      this.validationIssues.push(...result.issues);
      this.errorRows.push(...result.errorRows);
      this.diffPreview.push(...result.diffs);

      Object.entries(result.fillRates).forEach(([column, fillRate]) => {
        summary.fillRates[`${pair.target.name}.${column}`] = fillRate;
      });
    }

    const outputName = this.targetWorkbook.fileName.replace(/\.xlsx?$/i, '') || 'migrated';

    if (this.downloadBackup) {
      XLSX.writeFile(this.targetWorkbook.workbook, `${outputName}-backup.xlsx`);
    }

    XLSX.writeFile(destinationWorkbook, `${outputName}-migrated.xlsx`);

    this.summary = summary;
    this.auditLog.unshift({
      timestamp: new Date().toISOString(),
      sourceFile: this.sourceWorkbook.fileName,
      targetFile: this.targetWorkbook.fileName,
      batchMode: this.runBatchMode,
      processed: summary.processed,
      migrated: summary.migrated,
      failed: summary.failed
    });

    this.infoMessage = `Migration complete. Downloaded ${outputName}-migrated.xlsx`;
  }

  downloadErrorReport(): void {
    if (!this.errorRows.length || !this.targetWorkbook) {
      return;
    }

    const wb = XLSX.utils.book_new();
    const rows = this.errorRows.map((item) => ({
      Sheet: item.sheet,
      Row: item.row,
      Reason: item.reason,
      Source: JSON.stringify(item.sourceRow)
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, 'Migration Errors');

    const outputName = this.targetWorkbook.fileName.replace(/\.xlsx?$/i, '') || 'migrated';
    XLSX.writeFile(wb, `${outputName}-errors.xlsx`);
  }

  get sourceSheet(): SheetData | undefined {
    return this.sourceWorkbook?.sheets.find((sheet) => sheet.name === this.selectedSourceSheet);
  }

  get targetSheet(): SheetData | undefined {
    return this.targetWorkbook?.sheets.find((sheet) => sheet.name === this.selectedTargetSheet);
  }

  get readyForMapping(): boolean {
    return Boolean(this.sourceWorkbook && this.targetWorkbook);
  }

  private getSheetPairs(): { source: SheetData; target: SheetData }[] {
    if (!this.sourceWorkbook || !this.targetWorkbook || !this.sourceSheet || !this.targetSheet) {
      return [];
    }

    if (!this.runBatchMode) {
      return [{ source: this.sourceSheet, target: this.targetSheet }];
    }

    const sourceByName = new Map(this.sourceWorkbook.sheets.map((sheet) => [sheet.name, sheet]));
    return this.targetWorkbook.sheets
      .map((target) => ({ source: sourceByName.get(target.name), target }))
      .filter((pair): pair is { source: SheetData; target: SheetData } => Boolean(pair.source));
  }

  private migrateSheet(
    sourceSheet: SheetData,
    targetSheet: SheetData,
    mappings: ColumnMapping[],
    destinationWorkbook: XLSX.WorkBook
  ): SheetMigrationResult {
    const outputRows = targetSheet.rows.map((row) => ({ ...row }));
    const issues: ValidationIssue[] = [];
    const errors: ErrorRow[] = [];
    const diffs: DiffPreview[] = [];

    let processed = 0;
    let migrated = 0;
    let failed = 0;
    let skipped = 0;

    const useKeyMode =
      this.useKeyMergeMode &&
      sourceSheet.headers.includes(this.sourceKeyColumn) &&
      targetSheet.headers.includes(this.targetKeyColumn);

    if (this.useKeyMergeMode && !useKeyMode) {
      issues.push({
        sheet: targetSheet.name,
        row: 1,
        column: 'Key Merge',
        reason: 'Key merge enabled but selected key columns are missing in this sheet pair.'
      });
    }

    if (useKeyMode) {
      const targetKeyIndex = new Map<string, number>();
      outputRows.forEach((row, index) => {
        const key = this.toKey(row[this.targetKeyColumn]);
        if (key) {
          targetKeyIndex.set(key, index);
        }
      });

      const sourceSeenKeys = new Set<string>();

      sourceSheet.rows.forEach((sourceRow, sourceIndex) => {
        processed += 1;

        const sourceKeyRaw = sourceRow[this.sourceKeyColumn];
        const sourceKey = this.toKey(sourceKeyRaw);

        if (!sourceKey) {
          const reason = `Missing source key '${this.sourceKeyColumn}'`;
          issues.push({
            sheet: targetSheet.name,
            row: sourceIndex + 2,
            column: this.sourceKeyColumn,
            reason
          });
          errors.push({ sheet: targetSheet.name, row: sourceIndex + 2, reason, sourceRow });
          failed += 1;
          return;
        }

        if (sourceSeenKeys.has(sourceKey)) {
          const reason = `Duplicate source key '${sourceKey}'`;
          issues.push({
            sheet: targetSheet.name,
            row: sourceIndex + 2,
            column: this.sourceKeyColumn,
            reason
          });
          errors.push({ sheet: targetSheet.name, row: sourceIndex + 2, reason, sourceRow });
          failed += 1;
          return;
        }

        sourceSeenKeys.add(sourceKey);

        const existingIndex = targetKeyIndex.get(sourceKey);
        const isNewRecord = existingIndex === undefined;
        const rowIndex = isNewRecord ? outputRows.length : existingIndex;
        const baseRecord = isNewRecord
          ? this.createEmptyRecord(targetSheet.headers)
          : { ...outputRows[rowIndex] };

        baseRecord[this.targetKeyColumn] = sourceKeyRaw;
        const before = { ...baseRecord };

        const rowResult = this.applyMappingsToRecord(
          sourceRow,
          baseRecord,
          mappings,
          sourceIndex + 2,
          targetSheet.name
        );

        if (rowResult.failed) {
          failed += 1;
          issues.push(...rowResult.issues);
          errors.push({
            sheet: targetSheet.name,
            row: sourceIndex + 2,
            reason: rowResult.issues.map((item) => item.reason).join('; '),
            sourceRow
          });
          return;
        }

        if (isNewRecord) {
          outputRows.push(rowResult.record);
          targetKeyIndex.set(sourceKey, outputRows.length - 1);
        } else {
          outputRows[rowIndex] = rowResult.record;
        }

        migrated += 1;
        const changes = this.collectChanges(before, rowResult.record, targetSheet.headers);
        if (changes.length && diffs.length < 25) {
          diffs.push({ sheet: targetSheet.name, row: sourceIndex + 2, changes });
        }
      });
    } else {
      const maxRowCount = Math.max(sourceSheet.rows.length, outputRows.length);

      for (let rowIndex = 0; rowIndex < maxRowCount; rowIndex++) {
        const sourceRow = sourceSheet.rows[rowIndex];
        const existing = outputRows[rowIndex] ?? this.createEmptyRecord(targetSheet.headers);

        if (!sourceRow) {
          skipped += 1;
          if (rowIndex >= outputRows.length) {
            outputRows.push(existing);
          }
          continue;
        }

        processed += 1;
        const before = { ...existing };

        const rowResult = this.applyMappingsToRecord(
          sourceRow,
          { ...existing },
          mappings,
          rowIndex + 2,
          targetSheet.name
        );

        if (rowResult.failed) {
          failed += 1;
          issues.push(...rowResult.issues);
          errors.push({
            sheet: targetSheet.name,
            row: rowIndex + 2,
            reason: rowResult.issues.map((item) => item.reason).join('; '),
            sourceRow
          });
          continue;
        }

        if (rowIndex >= outputRows.length) {
          outputRows.push(rowResult.record);
        } else {
          outputRows[rowIndex] = rowResult.record;
        }

        migrated += 1;
        const changes = this.collectChanges(before, rowResult.record, targetSheet.headers);
        if (changes.length && diffs.length < 25) {
          diffs.push({ sheet: targetSheet.name, row: rowIndex + 2, changes });
        }
      }
    }

    const matrix: unknown[][] = [targetSheet.headers];
    outputRows.forEach((row) => {
      matrix.push(targetSheet.headers.map((header) => row[header] ?? ''));
    });

    destinationWorkbook.Sheets[targetSheet.name] = XLSX.utils.aoa_to_sheet(matrix);

    const fillRates: Record<string, number> = {};
    mappings.forEach((mapping) => {
      const nonEmpty = outputRows.filter((row) => !this.isEmptyValue(row[mapping.targetColumn])).length;
      fillRates[mapping.targetColumn] = outputRows.length
        ? Math.round((nonEmpty / outputRows.length) * 100)
        : 0;
    });

    return {
      processed,
      migrated,
      failed,
      skipped,
      issues,
      errorRows: errors,
      diffs,
      fillRates
    };
  }

  private applyMappingsToRecord(
    sourceRow: Record<string, unknown>,
    targetRecord: Record<string, unknown>,
    mappings: ColumnMapping[],
    sourceRowNumber: number,
    sheetName: string
  ): { record: Record<string, unknown>; issues: ValidationIssue[]; failed: boolean } {
    const issues: ValidationIssue[] = [];

    mappings.forEach((mapping) => {
      let value = this.resolveMappingValue(sourceRow, mapping);
      value = this.applyTransform(value, mapping.transformRule);

      const coercion = this.coerceType(value, mapping.targetType);
      if (!coercion.ok) {
        issues.push({
          sheet: sheetName,
          row: sourceRowNumber,
          column: mapping.targetColumn,
          reason: coercion.reason
        });
        return;
      }

      targetRecord[mapping.targetColumn] = coercion.value;

      if (mapping.required && this.isEmptyValue(targetRecord[mapping.targetColumn])) {
        issues.push({
          sheet: sheetName,
          row: sourceRowNumber,
          column: mapping.targetColumn,
          reason: 'Required value is empty'
        });
      }
    });

    return {
      record: targetRecord,
      issues,
      failed: issues.length > 0
    };
  }

  private resolveMappingValue(sourceRow: Record<string, unknown>, mapping: ColumnMapping): unknown {
    const sourceValue = mapping.sourceColumn ? sourceRow[mapping.sourceColumn] : '';
    const defaultValue = mapping.defaultValue;

    if (mapping.fallbackMode === 'always') {
      return defaultValue;
    }

    if (mapping.fallbackMode === 'when_source_equals') {
      const sourceAsString = String(sourceValue ?? '').trim();
      return sourceAsString === mapping.fallbackCompareValue ? defaultValue : sourceValue;
    }

    return this.isEmptyValue(sourceValue) ? defaultValue : sourceValue;
  }

  private applyTransform(value: unknown, rule: TransformRule): unknown {
    if (value === null || value === undefined) {
      return '';
    }

    if (rule === 'none') {
      return value;
    }

    const asString = String(value);

    if (rule === 'trim') {
      return asString.trim();
    }

    if (rule === 'uppercase') {
      return asString.toUpperCase();
    }

    if (rule === 'lowercase') {
      return asString.toLowerCase();
    }

    if (rule === 'number') {
      const parsed = Number(asString);
      return Number.isFinite(parsed) ? parsed : value;
    }

    const date = new Date(asString);
    if (Number.isNaN(date.getTime())) {
      return value;
    }

    return date.toISOString().slice(0, 10);
  }

  private coerceType(value: unknown, targetType: TargetType): { ok: true; value: unknown } | { ok: false; reason: string } {
    if (targetType === 'any') {
      return { ok: true, value };
    }

    if (targetType === 'string') {
      return { ok: true, value: String(value ?? '') };
    }

    if (targetType === 'number') {
      if (this.isEmptyValue(value)) {
        return { ok: true, value: '' };
      }

      const parsed = Number(value);
      if (!Number.isFinite(parsed)) {
        return { ok: false, reason: `Value '${String(value)}' is not a valid number` };
      }

      return { ok: true, value: parsed };
    }

    if (targetType === 'boolean') {
      if (this.isEmptyValue(value)) {
        return { ok: true, value: '' };
      }

      const normalized = String(value).trim().toLowerCase();
      if (['true', '1', 'yes', 'y'].includes(normalized)) {
        return { ok: true, value: true };
      }

      if (['false', '0', 'no', 'n'].includes(normalized)) {
        return { ok: true, value: false };
      }

      return { ok: false, reason: `Value '${String(value)}' is not a valid boolean` };
    }

    if (this.isEmptyValue(value)) {
      return { ok: true, value: '' };
    }

    const date = new Date(String(value));
    if (Number.isNaN(date.getTime())) {
      return { ok: false, reason: `Value '${String(value)}' is not a valid date` };
    }

    return { ok: true, value: date.toISOString().slice(0, 10) };
  }

  private collectChanges(
    before: Record<string, unknown>,
    after: Record<string, unknown>,
    headers: string[]
  ): string[] {
    return headers
      .filter((header) => String(before[header] ?? '') !== String(after[header] ?? ''))
      .map((header) => `${header}: '${String(before[header] ?? '')}' -> '${String(after[header] ?? '')}'`)
      .slice(0, 8);
  }

  private createEmptyRecord(headers: string[]): Record<string, unknown> {
    const record: Record<string, unknown> = {};
    headers.forEach((header) => {
      record[header] = '';
    });
    return record;
  }

  private buildAutoMappings(sourceHeaders: string[], targetHeaders: string[]): ColumnMapping[] {
    const normalizedSourceHeaders = new Map(
      sourceHeaders.map((header) => [header.trim().toLowerCase(), header])
    );

    return targetHeaders.map((targetColumn) => {
      const sourceColumn = normalizedSourceHeaders.get(targetColumn.trim().toLowerCase()) ?? '';
      return {
        targetColumn,
        sourceColumn,
        defaultValue: '',
        fallbackMode: 'when_source_empty',
        fallbackCompareValue: '',
        transformRule: 'none',
        targetType: 'any',
        required: false
      };
    });
  }

  private isEmptyValue(value: unknown): boolean {
    return value === null || value === undefined || String(value).trim() === '';
  }

  private toKey(value: unknown): string {
    return String(value ?? '').trim();
  }

  private extractFile(event: Event): File | null {
    const input = event.target as HTMLInputElement;
    return input.files?.item(0) ?? null;
  }

  private async parseWorkbook(file: File): Promise<WorkbookData> {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });

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
      workbook,
      sheets
    };
  }

  private toSheetData(
    name: string,
    matrix: (string | number | boolean | null)[][]
  ): SheetData {
    const headerRowIndex = matrix.findIndex((row) => row.some((cell) => String(cell ?? '').trim() !== ''));

    if (headerRowIndex < 0) {
      return { name, headers: [], rows: [] };
    }

    const rawHeaders = matrix[headerRowIndex] ?? [];
    const headers = this.normalizeHeaders(rawHeaders);

    const rows = matrix
      .slice(headerRowIndex + 1)
      .filter((row) => row.some((cell) => String(cell ?? '').trim() !== ''))
      .map((row) => {
        const record: Record<string, unknown> = {};
        headers.forEach((header, index) => {
          record[header] = row[index] ?? '';
        });
        return record;
      });

    return { name, headers, rows };
  }

  private normalizeHeaders(values: (string | number | boolean | null)[]): string[] {
    const seen = new Map<string, number>();

    return values.map((value, index) => {
      const base = String(value ?? '').trim() || `Column ${index + 1}`;
      const count = seen.get(base) ?? 0;
      seen.set(base, count + 1);

      return count === 0 ? base : `${base}_${count + 1}`;
    });
  }

  private rebuildMappings(): void {
    const sourceHeaders = this.sourceSheet?.headers ?? [];
    const targetHeaders = this.targetSheet?.headers ?? [];
    const sourceHeaderSet = new Set(sourceHeaders);
    const normalizedSourceHeaders = new Map(
      sourceHeaders.map((header) => [header.trim().toLowerCase(), header])
    );

    const previousMappings = new Map(this.mappings.map((mapping) => [mapping.targetColumn, mapping]));

    this.mappings = targetHeaders.map((targetColumn) => {
      const previous = previousMappings.get(targetColumn);
      const previousSource = previous?.sourceColumn ?? '';
      const autoMatchedSource = normalizedSourceHeaders.get(targetColumn.trim().toLowerCase()) ?? '';

      return {
        targetColumn,
        sourceColumn: sourceHeaderSet.has(previousSource) ? previousSource : autoMatchedSource,
        defaultValue: previous?.defaultValue ?? '',
        fallbackMode: previous?.fallbackMode ?? 'when_source_empty',
        fallbackCompareValue: previous?.fallbackCompareValue ?? '',
        transformRule: previous?.transformRule ?? 'none',
        targetType: previous?.targetType ?? 'any',
        required: previous?.required ?? false
      };
    });
  }

  private clearMessages(): void {
    this.errorMessage = '';
    this.infoMessage = '';
  }

  private toErrorMessage(error: unknown, fallback: string): string {
    if (error instanceof Error && error.message) {
      return error.message;
    }
    return fallback;
  }
}
