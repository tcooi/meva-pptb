import writeXlsxFile from 'write-excel-file/browser';
import { effectiveValue, type CompareStatus, type EnvironmentSnapshot, type VariableRow } from './environmentVariables';

export type ExportFormat = 'xlsx' | 'csv' | 'json';

export const EXPORT_FORMATS: { format: ExportFormat; label: string; filter: ToolBoxAPI.FileDialogFilter }[] = [
    { format: 'xlsx', label: 'Excel (.xlsx)', filter: { name: 'Excel workbook', extensions: ['xlsx'] } },
    { format: 'csv', label: 'CSV (.csv)', filter: { name: 'CSV', extensions: ['csv'] } },
    { format: 'json', label: 'JSON (.json)', filter: { name: 'JSON', extensions: ['json'] } },
];

// Same labels as the grid's comparison pills
const STATUS_LABELS: Record<CompareStatus, string> = { single: '', match: 'Same', different: 'Different', missing: 'Missing' };

interface ExportTable {
    headers: string[];
    data: (string | null)[][];
}

/** One row per variable with the same columns as the grid; shared by the CSV and Excel exports. */
function buildExportTable(rows: VariableRow[], snapshots: EnvironmentSnapshot[]): ExportTable {
    const isComparing = snapshots.length > 1;
    const headers = ['Display name', 'Schema name', 'Type', 'Description', 'Solutions'];
    if (isComparing) {
        headers.push('Comparison', ...snapshots.flatMap((s) => [`${s.connection.name} – Current value`, `${s.connection.name} – Default value`]));
    } else {
        headers.push('Default value', 'Current value');
    }

    const data = rows.map((row) => {
        const cells: (string | null)[] = [row.displayName, row.schemaName, row.type, row.description, row.solutions.join('; ')];
        if (isComparing) {
            cells.push(STATUS_LABELS[row.status]);
            row.values.forEach((variable, i) => {
                if (snapshots[i].error) cells.push('Failed to load', 'Failed to load');
                else if (!variable) cells.push('Not in this environment', 'Not in this environment');
                else cells.push(variable.currentValue, variable.defaultValue);
            });
        } else {
            const single = row.values[0];
            cells.push(single?.defaultValue ?? null, single?.currentValue ?? null);
        }
        return cells;
    });

    return { headers, data };
}

function csvField(value: string | null): string {
    if (value === null) return '';
    return /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

/** RFC 4180 CSV with a UTF-8 BOM so Excel reads non-ASCII characters correctly. */
function toCsv({ headers, data }: ExportTable): string {
    return '﻿' + [headers, ...data].map((cells) => cells.map(csvField).join(',')).join('\r\n') + '\r\n';
}

/** Structured export that keeps "not set" (null) and "empty" ("") apart. */
function toJson(rows: VariableRow[], snapshots: EnvironmentSnapshot[]): string {
    const isComparing = snapshots.length > 1;
    return JSON.stringify(
        {
            exportedAt: new Date().toISOString(),
            environments: snapshots.map((s) => ({ name: s.connection.name, environment: s.connection.environment, url: s.connection.url })),
            variables: rows.map((row) => ({
                schemaName: row.schemaName,
                displayName: row.displayName,
                description: row.description,
                type: row.type,
                solutions: row.solutions,
                ...(isComparing && { comparison: STATUS_LABELS[row.status] }),
                values: Object.fromEntries(
                    row.values.map((variable, i) => [
                        snapshots[i].connection.name,
                        snapshots[i].error
                            ? { error: snapshots[i].error }
                            : variable && { defaultValue: variable.defaultValue, currentValue: variable.currentValue, effectiveValue: effectiveValue(variable) },
                    ])
                ),
            })),
        },
        null,
        2
    );
}

async function toXlsx({ headers, data }: ExportTable): Promise<Uint8Array> {
    const blob = await writeXlsxFile([headers.map((value) => ({ value, fontWeight: 'bold' as const })), ...data], {
        sheet: 'Environment variables',
        stickyRowsCount: 1,
        columns: headers.map((_h, i) => ({ width: i < 2 ? 36 : i === 2 ? 12 : 40 })),
    }).toBlob();
    return new Uint8Array(await blob.arrayBuffer());
}

export function exportFileName(snapshots: EnvironmentSnapshot[], format: ExportFormat): string {
    const scope = snapshots.length === 1 ? snapshots[0].connection.name : 'compare';
    const date = new Date().toISOString().slice(0, 10);
    const safe = scope.replace(/[\\/:*?"<>|\s]+/g, '-').replace(/^-+|-+$/g, '');
    return `environment-variables-${safe}-${date}.${format}`;
}

export async function buildExport(rows: VariableRow[], snapshots: EnvironmentSnapshot[], format: ExportFormat): Promise<string | Uint8Array> {
    switch (format) {
        case 'csv':
            return toCsv(buildExportTable(rows, snapshots));
        case 'json':
            return toJson(rows, snapshots);
        case 'xlsx':
            return toXlsx(buildExportTable(rows, snapshots));
    }
}
