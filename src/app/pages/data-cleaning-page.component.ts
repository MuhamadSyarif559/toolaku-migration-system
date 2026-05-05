import { CommonModule } from '@angular/common';
import { Component, OnInit } from '@angular/core';
import { FormsModule } from '@angular/forms';

const CLEANING_API_BASE_URL = 'http://localhost:3001/api/data-cleaning';

interface CleanerInput {
  key: string;
  label: string;
  hint: string;
}

interface CleanerTool {
  mode: string;
  category: string;
  title: string;
  description: string;
  outputName: string;
  requiredFields: string[];
  inputs: CleanerInput[];
}

interface UploadPayloadFile {
  fieldName: string;
  fileName: string;
  contentBase64: string;
}

interface CleaningCatalogResponse {
  ok: boolean;
  tools: CleanerTool[];
}

interface CleaningApiResponse {
  ok: boolean;
  message: string;
  mode?: string;
  downloadFileName?: string;
  outputBase64?: string;
  stdout?: string;
  stderr?: string;
}

interface ToolRunState {
  files: Record<string, File | null>;
  infoMessage: string;
  errorMessage: string;
  stdout: string;
  outputFileName: string;
}

@Component({
  selector: 'app-data-cleaning-page',
  imports: [CommonModule, FormsModule],
  templateUrl: './data-cleaning-page.component.html',
  styleUrl: './data-cleaning-page.component.css'
})
export class DataCleaningPageComponent implements OnInit {
  readonly serverHint =
    'This feature runs local Python cleaners through the Node API. Start it with `npm run api:start` before using these tools.';

  tools: CleanerTool[] = [];
  toolStates: Record<string, ToolRunState> = {};

  catalogBusy = false;
  catalogErrorMessage = '';
  busyMode: string | null = null;

  async ngOnInit(): Promise<void> {
    await this.loadCatalog();
  }

  async loadCatalog(): Promise<void> {
    this.catalogBusy = true;
    this.catalogErrorMessage = '';

    try {
      const response = await fetch(`${CLEANING_API_BASE_URL}/catalog`);
      const data = (await response.json()) as CleaningCatalogResponse;

      if (!response.ok || !data.ok) {
        throw new Error('Could not load data cleaning catalog.');
      }

      this.tools = data.tools;
      this.toolStates = Object.fromEntries(
        this.tools.map((tool) => [
          tool.mode,
          {
            files: Object.fromEntries(tool.inputs.map((input) => [input.key, null])),
            infoMessage: '',
            errorMessage: '',
            stdout: '',
            outputFileName: tool.outputName
          }
        ])
      );
    } catch (error) {
      this.catalogErrorMessage = this.toErrorMessage(
        error,
        'Could not load cleaners. Check that the local API is running.'
      );
    } finally {
      this.catalogBusy = false;
    }
  }

  get groupedTools(): Array<{ category: string; tools: CleanerTool[] }> {
    const groups = new Map<string, CleanerTool[]>();
    for (const tool of this.tools) {
      const existing = groups.get(tool.category) ?? [];
      existing.push(tool);
      groups.set(tool.category, existing);
    }

    return Array.from(groups.entries()).map(([category, tools]) => ({ category, tools }));
  }

  getState(mode: string): ToolRunState {
    return this.toolStates[mode];
  }

  isBusy(mode: string): boolean {
    return this.busyMode === mode;
  }

  onFileSelected(mode: string, field: string, event: Event): void {
    const state = this.getState(mode);
    state.files[field] = this.extractFile(event);
    state.infoMessage = '';
    state.errorMessage = '';
    state.stdout = '';
  }

  hasSelectedFiles(mode: string): boolean {
    const state = this.getState(mode);
    return Object.values(state.files).some((file) => Boolean(file));
  }

  async runTool(tool: CleanerTool): Promise<void> {
    const state = this.getState(tool.mode);
    state.infoMessage = '';
    state.errorMessage = '';
    state.stdout = '';

    const selectedFiles = Object.entries(state.files).filter((entry): entry is [string, File] => Boolean(entry[1]));

    for (const requiredField of tool.requiredFields) {
      if (!state.files[requiredField]) {
        const missingInput = tool.inputs.find((input) => input.key === requiredField);
        state.errorMessage = `Missing required file: ${missingInput?.label || requiredField}`;
        return;
      }
    }

    if (!selectedFiles.length) {
      state.errorMessage = 'Upload the required workbook files before running this step.';
      return;
    }

    try {
      this.busyMode = tool.mode;
      const payloadFiles = await Promise.all(
        selectedFiles.map(async ([fieldName, file]) => ({
          fieldName,
          fileName: file.name,
          contentBase64: await this.toBase64(file)
        }))
      );

      const response = await fetch(`${CLEANING_API_BASE_URL}/run`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          mode: tool.mode,
          outputFileName: state.outputFileName.trim() || tool.outputName,
          files: payloadFiles
        })
      });

      const data = (await response.json()) as CleaningApiResponse;
      state.stdout = [data.stdout?.trim(), data.stderr?.trim()].filter(Boolean).join('\n');

      if (!response.ok || !data.ok || !data.outputBase64 || !data.downloadFileName) {
        throw new Error(data.message || 'Data cleaning failed.');
      }

      this.downloadBase64File(data.outputBase64, data.downloadFileName);
      state.infoMessage = `${data.message} Downloaded ${data.downloadFileName}`;
    } catch (error) {
      state.errorMessage = this.toErrorMessage(
        error,
        'Could not run the cleaning step. Check that the local API is running.'
      );
    } finally {
      this.busyMode = null;
    }
  }

  private extractFile(event: Event): File | null {
    const input = event.target as HTMLInputElement;
    return input.files?.item(0) ?? null;
  }

  private async toBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result;
        if (typeof result !== 'string') {
          reject(new Error('Could not encode the selected file.'));
          return;
        }
        const commaIndex = result.indexOf(',');
        resolve(commaIndex >= 0 ? result.slice(commaIndex + 1) : result);
      };
      reader.onerror = () => reject(reader.error ?? new Error('Could not read the selected file.'));
      reader.readAsDataURL(file);
    });
  }

  private downloadBase64File(contentBase64: string, fileName: string): void {
    const byteCharacters = atob(contentBase64);
    const byteNumbers = new Array<number>(byteCharacters.length);
    for (let index = 0; index < byteCharacters.length; index++) {
      byteNumbers[index] = byteCharacters.charCodeAt(index);
    }

    const blob = new Blob([new Uint8Array(byteNumbers)], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = objectUrl;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(objectUrl);
  }

  private toErrorMessage(error: unknown, fallback: string): string {
    if (error instanceof Error && error.message) {
      return error.message;
    }
    return fallback;
  }
}
