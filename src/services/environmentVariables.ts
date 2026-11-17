export type VariableType = 'String' | 'Number' | 'Boolean' | 'JSON' | 'Data source' | 'Secret' | 'Unknown';

export interface EnvironmentVariable {
    definitionId: string;
    schemaName: string;
    displayName: string;
    description: string;
    type: VariableType;
    isManaged: boolean;
    defaultValue: string | null;
    currentValue: string | null;
    /** Id of the environmentvariablevalue record holding the current value, if there is one. */
    currentValueId: string | null;
}

export interface EnvironmentSnapshot {
    /** Zero-based connection slot this snapshot was loaded from. */
    slot: number;
    connection: ToolBoxAPI.Connection;
    /** Keyed by lower-cased schema name so environments can be matched up. */
    variables: Map<string, EnvironmentVariable>;
    error?: string;
}

/** Which of a variable's two values is being edited. */
export type ValueKind = 'current' | 'default';

/** A value to open straight in edit mode: the environment (connection slot) and which value. */
export interface EditTarget {
    slot: number;
    kind: ValueKind;
}

export type CompareStatus = 'single' | 'match' | 'different' | 'missing';

export interface VariableRow {
    key: string;
    schemaName: string;
    displayName: string;
    description: string;
    type: VariableType;
    /** One entry per snapshot, in snapshot order; null when the variable doesn't exist there. */
    values: (EnvironmentVariable | null)[];
    solutions: string[];
    status: CompareStatus;
}

export interface SolutionInfo {
    uniqueName: string;
    friendlyName: string;
}

const TYPE_BY_OPTION: Record<number, VariableType> = {
    100000000: 'String',
    100000001: 'Number',
    100000002: 'Boolean',
    100000003: 'JSON',
    100000004: 'Data source',
    100000005: 'Secret',
};

// Solutions that contain (almost) every component and so are useless as a filter
const HIDDEN_SOLUTIONS = new Set(['active', 'default']);

/**
 * Runs an OData query and follows @odata.nextLink so large environments aren't truncated.
 */
async function queryAll(odataQuery: string, target: ToolBoxAPI.ConnectionTarget): Promise<Record<string, unknown>[]> {
    const records: Record<string, unknown>[] = [];
    let next: string | undefined = odataQuery;

    while (next) {
        const result = (await window.dataverseAPI.queryData(next, target)) as { value: Record<string, unknown>[]; '@odata.nextLink'?: string };
        records.push(...result.value);
        // nextLink is absolute; queryData expects a path relative to the Web API root
        const nextLink = result['@odata.nextLink'];
        next = nextLink ? nextLink.replace(/^.*\/api\/data\/v[\d.]+\//, '') : undefined;
    }

    return records;
}

const DEFINITION_QUERY =
    'environmentvariabledefinitions' +
    '?$select=environmentvariabledefinitionid,schemaname,displayname,description,type,defaultvalue,ismanaged' +
    '&$expand=environmentvariabledefinition_environmentvariablevalue($select=environmentvariablevalueid,value)';

function toVariable(record: Record<string, unknown>): EnvironmentVariable {
    const valueRecords = (record.environmentvariabledefinition_environmentvariablevalue as { environmentvariablevalueid?: string; value?: string | null }[] | undefined) ?? [];
    const schemaName = String(record.schemaname ?? '');
    return {
        definitionId: String(record.environmentvariabledefinitionid),
        schemaName,
        displayName: String(record.displayname ?? schemaName),
        description: String(record.description ?? ''),
        type: TYPE_BY_OPTION[record.type as number] ?? 'Unknown',
        isManaged: Boolean(record.ismanaged),
        defaultValue: (record.defaultvalue as string | null | undefined) ?? null,
        // Dataverse allows at most one current value per definition
        currentValue: valueRecords[0]?.value ?? null,
        currentValueId: valueRecords[0]?.environmentvariablevalueid ?? null,
    };
}

export async function loadEnvironmentVariables(target: ToolBoxAPI.ConnectionTarget): Promise<Map<string, EnvironmentVariable>> {
    const variables = new Map<string, EnvironmentVariable>();
    for (const record of await queryAll(DEFINITION_QUERY, target)) {
        const variable = toVariable(record);
        variables.set(variable.schemaName.toLowerCase(), variable);
    }
    return variables;
}

/** Re-reads one definition after an edit, so only that variable needs refreshing. */
export async function loadEnvironmentVariable(target: ToolBoxAPI.ConnectionTarget, definitionId: string): Promise<EnvironmentVariable | null> {
    const records = await queryAll(`${DEFINITION_QUERY}&$filter=environmentvariabledefinitionid eq ${definitionId}`, target);
    return records.length > 0 ? toVariable(records[0]) : null;
}

/** Sets the current value: updates the existing value record, or creates one linked to the definition. */
export async function saveCurrentValue(target: ToolBoxAPI.ConnectionTarget, variable: EnvironmentVariable, value: string): Promise<void> {
    if (variable.currentValueId) {
        await window.dataverseAPI.update('environmentvariablevalue', variable.currentValueId, { value }, target);
    } else {
        await window.dataverseAPI.create(
            'environmentvariablevalue',
            {
                schemaname: variable.schemaName,
                value,
                'EnvironmentVariableDefinitionId@odata.bind': `/environmentvariabledefinitions(${variable.definitionId})`,
            },
            target
        );
    }
}

/** Deletes the current value record so the default value applies again. */
export async function removeCurrentValue(target: ToolBoxAPI.ConnectionTarget, variable: EnvironmentVariable): Promise<void> {
    if (variable.currentValueId) {
        await window.dataverseAPI.delete('environmentvariablevalue', variable.currentValueId, target);
    }
}

/** Sets (or with null, clears) the default value on the definition. */
export async function saveDefaultValue(target: ToolBoxAPI.ConnectionTarget, variable: EnvironmentVariable, value: string | null): Promise<void> {
    await window.dataverseAPI.update('environmentvariabledefinition', variable.definitionId, { defaultvalue: value }, target);
}

const OPTION_BY_TYPE: Partial<Record<VariableType, number>> = {
    String: 100000000,
    Number: 100000001,
    Boolean: 100000002,
    JSON: 100000003,
};

/** Types a new variable can be created as here; Data source and Secret need extra setup in the maker portal. */
export const CREATABLE_TYPES = Object.keys(OPTION_BY_TYPE) as VariableType[];

export interface UnmanagedSolution extends SolutionInfo {
    /** The solution publisher's customization prefix, used for new schema names (e.g. "contoso"). */
    prefix: string;
}

/** Unmanaged, visible solutions a new variable can be added to, with their publisher prefix. */
export async function loadUnmanagedSolutions(target: ToolBoxAPI.ConnectionTarget): Promise<UnmanagedSolution[]> {
    const records = await queryAll(
        'solutions?$select=uniquename,friendlyname&$filter=isvisible eq true and ismanaged eq false' +
            '&$expand=publisherid($select=customizationprefix)&$orderby=friendlyname',
        target
    );
    return records
        .filter((r) => !HIDDEN_SOLUTIONS.has(String(r.uniquename).toLowerCase()))
        .map((r) => ({
            uniqueName: String(r.uniquename),
            friendlyName: String(r.friendlyname || r.uniquename),
            prefix: String((r.publisherid as { customizationprefix?: string } | undefined)?.customizationprefix ?? ''),
        }));
}

/** Suggests the part of a schema name after the prefix, e.g. "API base URL" -> "APIBaseURL". */
export function suggestSchemaName(displayName: string): string {
    return displayName
        .split(/[^A-Za-z0-9]+/)
        .filter(Boolean)
        .map((word) => word[0].toUpperCase() + word.slice(1))
        .join('');
}

/** Returns a problem with a full schema name (prefix_Name), or undefined when it's valid. */
export function validateSchemaName(schemaName: string, existing: Map<string, EnvironmentVariable>): string | undefined {
    if (!/^[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+$/.test(schemaName)) {
        return 'Use a prefix, an underscore and then letters, numbers or underscores (e.g. contoso_ApiUrl).';
    }
    if (schemaName.length > 100) {
        return 'Schema names can be at most 100 characters.';
    }
    if (existing.has(schemaName.toLowerCase())) {
        return 'A variable with this schema name already exists in this environment.';
    }
    return undefined;
}

export interface NewVariable {
    schemaName: string;
    displayName: string;
    description: string;
    type: VariableType;
    defaultValue: string | null;
    currentValue: string | null;
    /** Unmanaged solution to add the definition to, if any. */
    solutionUniqueName: string | null;
}

/**
 * Creates a definition (with its default value), adds it to the chosen solution and sets the current value.
 * Steps after the definition exists are reported as warnings rather than errors, so the caller can still
 * show the new variable and tell the user exactly what's left to do.
 */
export async function createEnvironmentVariable(target: ToolBoxAPI.ConnectionTarget, input: NewVariable): Promise<{ definitionId: string; warnings: string[] }> {
    const { id: definitionId } = await window.dataverseAPI.create(
        'environmentvariabledefinition',
        {
            schemaname: input.schemaName,
            displayname: input.displayName,
            description: input.description || null,
            type: OPTION_BY_TYPE[input.type],
            defaultvalue: input.defaultValue,
        },
        target
    );

    const warnings: string[] = [];
    if (input.solutionUniqueName) {
        try {
            // AddSolutionComponent rather than the MSCRM.SolutionUniqueName header, which PPTB only sends after a consent prompt
            await window.dataverseAPI.execute(
                {
                    operationName: 'AddSolutionComponent',
                    operationType: 'action',
                    parameters: {
                        ComponentId: definitionId,
                        ComponentType: 380, // Environment Variable Definition
                        SolutionUniqueName: input.solutionUniqueName,
                        AddRequiredComponents: false,
                    },
                },
                target
            );
        } catch (error) {
            console.error(`Error adding ${input.schemaName} to solution ${input.solutionUniqueName}:`, error);
            warnings.push(`It wasn't added to the solution ${input.solutionUniqueName}: ${describeError(error)}`);
        }
    }

    if (input.currentValue !== null) {
        try {
            await window.dataverseAPI.create(
                'environmentvariablevalue',
                {
                    schemaname: input.schemaName,
                    value: input.currentValue,
                    'EnvironmentVariableDefinitionId@odata.bind': `/environmentvariabledefinitions(${definitionId})`,
                },
                target
            );
        } catch (error) {
            console.error(`Error setting the current value of ${input.schemaName}:`, error);
            warnings.push(`Its current value wasn't set: ${describeError(error)}`);
        }
    }

    return { definitionId, warnings };
}

/** Whether a variable can be copied to another environment; Data source and Secret need setup in the maker portal. */
export function canCopy(type: VariableType): boolean {
    return CREATABLE_TYPES.includes(type);
}

/**
 * Creates a variable from another environment's definition, keeping its schema name so the two match up.
 * The current value is usually environment-specific, so it's only copied when asked for.
 */
export async function copyEnvironmentVariable(
    target: ToolBoxAPI.ConnectionTarget,
    source: EnvironmentVariable,
    options: { includeCurrentValue: boolean; solutionUniqueName: string | null }
): Promise<{ definitionId: string; warnings: string[] }> {
    return createEnvironmentVariable(target, {
        schemaName: source.schemaName,
        displayName: source.displayName,
        description: source.description,
        type: source.type,
        defaultValue: source.defaultValue,
        currentValue: options.includeCurrentValue ? source.currentValue : null,
        solutionUniqueName: options.solutionUniqueName,
    });
}

/** Boolean values are stored as "yes"/"no", the same as the Power Apps maker portal writes them. */
export const BOOLEAN_VALUES = ['yes', 'no'] as const;

/** Returns a problem with the value for this type, or undefined when it's valid. */
export function validateValue(type: VariableType, value: string): string | undefined {
    switch (type) {
        case 'Number':
            return value.trim() !== '' && Number.isFinite(Number(value)) ? undefined : 'Enter a number.';
        case 'Boolean':
            return (BOOLEAN_VALUES as readonly string[]).includes(value) ? undefined : 'Choose Yes or No.';
        case 'JSON':
            try {
                JSON.parse(value);
                return undefined;
            } catch (error) {
                return `Not valid JSON: ${(error as Error).message}`;
            }
        default:
            return undefined;
    }
}

export interface SolutionData {
    /** Every visible solution in the environment, whether or not it contains environment variables. */
    solutions: SolutionInfo[];
    /** Solutions each environment variable definition belongs to, keyed by definition id. */
    membership: Map<string, SolutionInfo[]>;
}

/**
 * Loads the environment's solutions and which environment variable definitions each contains.
 * Two flat queries joined client-side, rather than $expand from solutioncomponent.
 */
export async function loadSolutionData(target: ToolBoxAPI.ConnectionTarget): Promise<SolutionData> {
    const [solutionRecords, componentRecords] = await Promise.all([
        queryAll('solutions?$select=solutionid,uniquename,friendlyname&$filter=isvisible eq true', target),
        // componenttype 380 = Environment Variable Definition
        queryAll('solutioncomponents?$select=objectid,_solutionid_value&$filter=componenttype eq 380', target),
    ]);

    const solutionsById = new Map<string, SolutionInfo>();
    for (const record of solutionRecords) {
        const uniqueName = String(record.uniquename ?? '');
        if (!uniqueName || HIDDEN_SOLUTIONS.has(uniqueName.toLowerCase())) {
            continue;
        }
        solutionsById.set(String(record.solutionid).toLowerCase(), { uniqueName, friendlyName: String(record.friendlyname || uniqueName) });
    }

    const membership = new Map<string, SolutionInfo[]>();
    for (const record of componentRecords) {
        const solution = solutionsById.get(String(record._solutionid_value ?? '').toLowerCase());
        if (!solution) {
            continue;
        }
        const definitionId = String(record.objectid);
        const list = membership.get(definitionId) ?? [];
        list.push(solution);
        membership.set(definitionId, list);
    }

    return { solutions: [...solutionsById.values()], membership };
}

/** The value Dataverse actually uses at runtime: current value if set, otherwise the default. */
export function effectiveValue(variable: EnvironmentVariable | null): string | null {
    if (!variable) {
        return null;
    }
    return variable.currentValue ?? variable.defaultValue;
}

function normalise(value: string | null, type: VariableType): string | null {
    if (value === null) {
        return null;
    }
    if (type === 'JSON') {
        // Ignore whitespace/formatting differences between environments
        try {
            return JSON.stringify(JSON.parse(value));
        } catch {
            return value.trim();
        }
    }
    if (type === 'Boolean') {
        return value.trim().toLowerCase();
    }
    return value.trim();
}

/**
 * Merges per-environment snapshots into one row per schema name.
 * Solution membership is taken from the first snapshot that has it (normally the primary connection).
 */
export function buildRows(snapshots: EnvironmentSnapshot[], membership: Map<string, SolutionInfo[]>): VariableRow[] {
    const loaded = snapshots.filter((s) => !s.error);
    const keys = new Set<string>();
    for (const snapshot of loaded) {
        snapshot.variables.forEach((_v, key) => keys.add(key));
    }

    const rows: VariableRow[] = [];
    for (const key of keys) {
        const values = snapshots.map((s) => (s.error ? null : s.variables.get(key) ?? null));
        const reference = values.find((v): v is EnvironmentVariable => v !== null)!;

        const solutions = new Set<string>();
        for (const value of values) {
            membership.get(value?.definitionId ?? '')?.forEach((s) => solutions.add(s.uniqueName));
        }

        let status: CompareStatus = 'single';
        if (loaded.length > 1) {
            const present = loaded.map((s) => s.variables.get(key) ?? null);
            if (present.some((v) => v === null)) {
                status = 'missing';
            } else {
                const distinct = new Set(present.map((v) => normalise(effectiveValue(v), reference.type)));
                status = distinct.size === 1 ? 'match' : 'different';
            }
        }

        rows.push({
            key,
            schemaName: reference.schemaName,
            displayName: reference.displayName,
            description: reference.description,
            type: reference.type,
            values,
            solutions: [...solutions],
            status,
        });
    }

    return rows.sort((a, b) => a.displayName.localeCompare(b.displayName, undefined, { sensitivity: 'base' }));
}

/** Maps Dataverse HTTP failures to a plain-language message for notifications. */
export function describeError(error: unknown): string {
    const status = (error as { status?: number })?.status;
    switch (true) {
        case status === 401:
            return 'Authentication failed, please reconnect.';
        case status === 403:
            return "You don't have permission to read this data in this environment.";
        case status === 429:
            return 'Too many requests, please wait and try again.';
        case status !== undefined && status >= 500:
            return 'Server error, try again later.';
        default:
            return (error as Error)?.message || 'Unknown error';
    }
}
