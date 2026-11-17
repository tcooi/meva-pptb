import { useCallback, useDeferredValue, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import {
    Button,
    Caption1,
    Popover,
    PopoverSurface,
    PopoverTrigger,
    Tooltip,
    FluentProvider,
    Menu,
    MenuItem,
    MenuList,
    MenuPopover,
    MenuTrigger,
    MessageBar,
    MessageBarBody,
    MessageBarTitle,
    Spinner,
    Text,
    makeStyles,
    mergeClasses,
    tokens,
    webDarkTheme,
    webLightTheme,
} from '@fluentui/react-components';
import { Add16Regular, ArrowClockwise16Regular, ArrowDownload16Regular, CopyArrowRight16Regular, ArrowDown12Regular, BracesVariable20Regular, DocumentBulletList24Regular, PlugDisconnected24Regular, Search24Regular } from '@fluentui/react-icons';
import { DEFAULT_FILTERS, FilterBar, type Filters } from './components/FilterBar';
import { CopyVariablesDialog } from './components/CopyVariablesDialog';
import { NewVariableDialog } from './components/NewVariableDialog';
import { VariableDetailsDialog } from './components/VariableDetailsDialog';
import { VariablesTable } from './components/VariablesTable';
import { AddEnvironmentHelp, EmptyState, EnvDot, MAX_ENVIRONMENTS } from './components/ui';
import { HIDDEN_VALUE, isFilterActive, matchesColumnFilter, type ColumnFilter } from './services/columnFilters';
import { notify, useConnections, useTheme, useToolboxEvents } from './hooks/useToolboxAPI';
import { buildExport, EXPORT_FORMATS, exportFileName, type ExportFormat } from './services/exportVariables';
import {
    buildRows,
    describeError,
    effectiveValue,
    loadEnvironmentVariable,
    loadEnvironmentVariables,
    loadSolutionData,
    type EditTarget,
    type EnvironmentSnapshot,
    type EnvironmentVariable,
    type SolutionInfo,
    type VariableRow,
} from './services/environmentVariables';

const useStyles = makeStyles({
    root: {
        display: 'flex',
        flexDirection: 'column',
        height: '100vh',
        backgroundColor: tokens.colorNeutralBackground1,
        overflow: 'hidden',
    },
    topBar: {
        display: 'flex',
        alignItems: 'center',
        gap: tokens.spacingHorizontalM,
        minHeight: '48px',
        padding: `0 ${tokens.spacingHorizontalL}`,
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
        flexWrap: 'wrap',
    },
    title: {
        display: 'flex',
        alignItems: 'center',
        gap: tokens.spacingHorizontalS,
        marginRight: 'auto',
        fontSize: tokens.fontSizeBase400,
        fontWeight: tokens.fontWeightSemibold,
    },
    titleIcon: {
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: '26px',
        height: '26px',
        borderRadius: tokens.borderRadiusMedium,
        backgroundColor: tokens.colorBrandBackground,
        color: tokens.colorNeutralForegroundOnBrand,
    },
    environments: {
        display: 'flex',
        flexWrap: 'wrap',
        alignItems: 'center',
        gap: tokens.spacingHorizontalL,
    },
    environment: {
        display: 'inline-flex',
        alignItems: 'center',
        gap: '6px',
        maxWidth: '240px',
        fontSize: tokens.fontSizeBase200,
        color: tokens.colorNeutralForeground2,
    },
    environmentError: {
        color: tokens.colorStatusDangerForeground1,
    },
    environmentName: {
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap',
        color: tokens.colorNeutralForeground1,
        fontWeight: tokens.fontWeightSemibold,
    },
    messages: {
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalXS,
        padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL}`,
    },
    footer: {
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        gap: tokens.spacingHorizontalM,
        minHeight: '30px',
        padding: `0 ${tokens.spacingHorizontalL}`,
        borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
        backgroundColor: tokens.colorNeutralBackground2,
        color: tokens.colorNeutralForeground3,
        fontVariantNumeric: 'tabular-nums',
    },
    addHint: {
        fontWeight: tokens.fontWeightSemibold,
        color: tokens.colorBrandForeground1,
    },
    addHintDetail: {
        display: 'inline-flex',
        alignItems: 'center',
        gap: '2px',
        marginLeft: tokens.spacingHorizontalXS,
        fontWeight: tokens.fontWeightRegular,
        color: tokens.colorNeutralForeground3,
    },
    addSurface: {
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalS,
        maxWidth: '300px',
        padding: tokens.spacingHorizontalM,
    },
    loadingScreen: {
        height: '100vh',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        gap: tokens.spacingVerticalM,
        backgroundColor: tokens.colorNeutralBackground1,
    },
    loadingDetail: {
        color: tokens.colorNeutralForeground3,
        textAlign: 'center',
        // Two lines reserved (as in index.html's boot loader) so the spinner stays put however long the text is
        minHeight: '40px',
        maxWidth: '480px',
        padding: `0 ${tokens.spacingHorizontalL}`,
    },
});

function matchesValueFilter(variable: EnvironmentVariable | null, filter: Filters['value']): boolean {
    if (!variable) return false;
    switch (filter) {
        case 'current':
            return variable.currentValue !== null;
        case 'default':
            return variable.currentValue === null && variable.defaultValue !== null;
        case 'none':
            return effectiveValue(variable) === null;
        default:
            return true;
    }
}

/**
 * Lower-cased value texts for a column cell. Secrets become HIDDEN_VALUE so they're never
 * matched as text; a variable missing from an environment contributes nothing.
 */
function valueTexts(variable: EnvironmentVariable | null, which: ('currentValue' | 'defaultValue')[]): string[] {
    if (!variable) return [];
    return which.map((key) => (variable.type === 'Secret' && variable[key] ? HIDDEN_VALUE : (variable[key] ?? '').toLowerCase()));
}

interface SearchIndex {
    /** Everything the main search box looks at. */
    all: string;
    /** Texts per grid column id, for the column filters. */
    columns: Record<string, string[]>;
}

/** Built once per load rather than per keystroke. */
function buildSearchIndex(row: VariableRow, snapshots: EnvironmentSnapshot[]): SearchIndex {
    const columns: Record<string, string[]> = {
        displayName: [row.displayName.toLowerCase()],
        schemaName: [row.schemaName.toLowerCase()],
        type: [row.type.toLowerCase()],
        comparison: [{ single: '', match: 'same', different: 'different', missing: 'missing' }[row.status]],
        defaultValue: row.values.flatMap((v) => valueTexts(v, ['defaultValue'])),
        currentValue: row.values.flatMap((v) => valueTexts(v, ['currentValue'])),
    };
    snapshots.forEach((snapshot, index) => {
        columns[`env${snapshot.slot}`] = valueTexts(row.values[index], ['currentValue', 'defaultValue']);
    });
    // The main search covers names, description and every value (type has a filter of its own)
    const values = [...columns.currentValue, ...columns.defaultValue].filter((t) => t !== HIDDEN_VALUE);
    const all = [columns.displayName[0], columns.schemaName[0], row.description.toLowerCase(), ...values].join('\n');
    return { all, columns };
}

function App({ initialTheme }: { initialTheme: 'light' | 'dark' }) {
    const styles = useStyles();
    const { theme, refreshTheme } = useTheme(initialTheme);
    const { connections, isLoading: isLoadingConnections, refreshConnections } = useConnections();
    const [snapshots, setSnapshots] = useState<EnvironmentSnapshot[]>([]);
    const [solutions, setSolutions] = useState<SolutionInfo[]>([]);
    const [membership, setMembership] = useState<Map<string, SolutionInfo[]>>(new Map());
    const [isLoading, setIsLoading] = useState(false);
    const [isLoadingSolutions, setIsLoadingSolutions] = useState(false);
    const [solutionsError, setSolutionsError] = useState<string | undefined>();
    // The full-screen loader stays up until the first set of variables arrives; later reloads keep the grid visible
    const [hasLoaded, setHasLoaded] = useState(false);
    const [filters, setFilters] = useState<Filters>(DEFAULT_FILTERS);
    // Keyed so the details dialog always shows the latest values after an edit
    const [selectedKey, setSelectedKey] = useState<string | null>(null);
    // Set when the dialog was opened from a cell's edit button, so it opens straight into that editor
    const [editTarget, setEditTarget] = useState<EditTarget | null>(null);
    const selectRow = useCallback((row: VariableRow, edit?: EditTarget) => {
        setSelectedKey(row.key);
        setEditTarget(edit ?? null);
    }, []);
    const setColumnFilter = useCallback((id: string, filter: ColumnFilter | undefined) => {
        setFilters((current) => {
            const columnFilters = { ...current.columnFilters };
            if (filter) columnFilters[id] = filter;
            else delete columnFilters[id];
            return { ...current, columnFilters };
        });
    }, []);
    // Ignore results from a load that has been superseded by a newer one
    const loadId = useRef(0);

    const loadAll = useCallback(async () => {
        const id = ++loadId.current;
        // Keep slot indexes: the index is the connection target, even with gaps
        const assigned = connections.map((connection, slot) => ({ connection, slot })).filter((c): c is { connection: ToolBoxAPI.Connection; slot: number } => !!c.connection);

        if (assigned.length === 0) {
            setSnapshots([]);
            return;
        }

        setIsLoading(true);
        setIsLoadingSolutions(true);

        // Solution data only feeds the solution filter and is the slower query, so it
        // runs alongside the variable queries and fills in after the grid is already showing.
        const solutionLoad = Promise.all(
            assigned.map(async ({ connection, slot }) => {
                try {
                    return await loadSolutionData(slot);
                } catch (error) {
                    console.error(`Error loading solutions from ${connection.name}:`, error);
                    return { error: `${connection.name}: ${describeError(error)}` };
                }
            })
        ).then((results) => {
            if (id !== loadId.current) return;

            const membershipByDefinition = new Map<string, SolutionInfo[]>();
            const uniqueSolutions = new Map<string, SolutionInfo>();
            const errors: string[] = [];
            for (const result of results) {
                if ('error' in result) {
                    errors.push(result.error);
                    continue;
                }
                result.solutions.forEach((solution) => uniqueSolutions.set(solution.uniqueName, solution));
                result.membership.forEach((list, definitionId) => membershipByDefinition.set(definitionId, list));
            }

            setMembership(membershipByDefinition);
            setSolutions([...uniqueSolutions.values()]);
            setSolutionsError(errors.length > 0 ? errors.join('\n') : undefined);
            // Shown on the Solution filter (warning icon + note) rather than as a PPTB notification on every load
            setIsLoadingSolutions(false);
        });

        try {
            const loaded = await Promise.all(
                assigned.map(async ({ connection, slot }): Promise<EnvironmentSnapshot> => {
                    try {
                        return { slot, connection, variables: await loadEnvironmentVariables(slot) };
                    } catch (error) {
                        console.error(`Error loading environment variables from ${connection.name}:`, error);
                        return { slot, connection, variables: new Map(), error: describeError(error) };
                    }
                })
            );

            if (id !== loadId.current) return;
            // Failed environments are shown as error bars above the grid rather than as PPTB notifications
            setSnapshots(loaded);
            setHasLoaded(true);
        } finally {
            if (id === loadId.current) setIsLoading(false);
        }

        await solutionLoad;
    }, [connections]);

    useEffect(() => {
        loadAll();
    }, [loadAll]);

    /** Re-reads a single variable in one environment after it's edited, instead of reloading everything. */
    const refreshVariable = useCallback(async (slot: number, definitionId: string) => {
        const variable = await loadEnvironmentVariable(slot, definitionId);
        setSnapshots((current) =>
            current.map((snapshot) => {
                if (snapshot.slot !== slot) return snapshot;
                const variables = new Map(snapshot.variables);
                for (const [key, existing] of variables) {
                    if (existing.definitionId === definitionId) variables.delete(key);
                }
                if (variable) variables.set(variable.schemaName.toLowerCase(), variable);
                return { ...snapshot, variables };
            })
        );
        return variable;
    }, []);

    const [isCreating, setIsCreating] = useState(false);

    /** Records new definitions as members of the solution they were added to, so the solution filter finds them. */
    const addToSolution = useCallback((definitionIds: string[], solution: SolutionInfo | null) => {
        if (!solution) return;
        const info = { uniqueName: solution.uniqueName, friendlyName: solution.friendlyName };
        setMembership((current) => {
            const next = new Map(current);
            definitionIds.forEach((id) => next.set(id, [info]));
            return next;
        });
        setSolutions((current) => (current.some((s) => s.uniqueName === info.uniqueName) ? current : [...current, info]));
    }, []);

    /** Shows a newly created variable: adds it to the grid and the solution filter, then opens its details. */
    const handleCreated = useCallback(
        async (slot: number, definitionId: string, solution: SolutionInfo | null) => {
            const variable = await refreshVariable(slot, definitionId);
            addToSolution([definitionId], solution);
            if (variable) {
                setEditTarget(null);
                setSelectedKey(variable.schemaName.toLowerCase());
            }
        },
        [refreshVariable, addToSolution]
    );

    // Rows (by key) waiting to be copied to another environment, and the environment picked for them if any
    const [copyRequest, setCopyRequest] = useState<{ keys: string[]; target?: number } | null>(null);
    const openCopy = useCallback((rows: VariableRow[], target?: number) => {
        setSelectedKey(null);
        setCopyRequest({ keys: rows.map((row) => row.key), target });
    }, []);
    const copySingle = useCallback((row: VariableRow, target: number) => openCopy([row], target), [openCopy]);

    /** Shows copied variables: one is re-read on its own, several reload that environment once. */
    const handleCopied = useCallback(
        async (slot: number, definitionIds: string[], solution: SolutionInfo | null) => {
            if (definitionIds.length === 1) {
                await refreshVariable(slot, definitionIds[0]);
            } else {
                const variables = await loadEnvironmentVariables(slot);
                setSnapshots((current) => current.map((snapshot) => (snapshot.slot === slot ? { ...snapshot, variables } : snapshot)));
            }
            addToSolution(definitionIds, solution);
        },
        [refreshVariable, addToSolution]
    );

    useToolboxEvents((event) => {
        if (event === 'connection:created' || event === 'connection:updated' || event === 'connection:deleted') {
            refreshConnections();
        } else if (event === 'settings:updated') {
            refreshTheme();
        }
    });

    const rows = useMemo(() => buildRows(snapshots, membership), [snapshots, membership]);
    const searchIndex = useMemo(() => new Map(rows.map((row) => [row.key, buildSearchIndex(row, snapshots)])), [rows, snapshots]);
    const isComparing = snapshots.length > 1;
    // Typing updates the input immediately; filtering the grid catches up without blocking keystrokes
    const deferredFilters = useDeferredValue(filters);

    const filteredRows = useMemo(() => {
        const filters = deferredFilters;
        const search = filters.search.trim().toLowerCase();
        const columnFilters = Object.entries(filters.columnFilters).filter((entry): entry is [string, ColumnFilter] => isFilterActive(entry[1]));
        return rows.filter(
            (row) =>
                (!search || searchIndex.get(row.key)!.all.includes(search)) &&
                columnFilters.every(([id, filter]) => {
                    const texts = searchIndex.get(row.key)!.columns[id];
                    // A filter for a column that's no longer shown (e.g. a removed environment) is ignored
                    return texts === undefined || matchesColumnFilter(texts, filter);
                }) &&
                (filters.type === 'all' || row.type === filters.type) &&
                (filters.solution === 'all' || row.solutions.includes(filters.solution)) &&
                (filters.value === 'all' || row.values.some((v) => matchesValueFilter(v, filters.value))) &&
                (!isComparing || filters.compare === 'all' || row.status === filters.compare)
        );
    }, [rows, searchIndex, deferredFilters, isComparing]);

    const failed = snapshots.filter((s) => s.error);
    const missingRows = useMemo(() => (isComparing ? filteredRows.filter((row) => row.status === 'missing') : []), [filteredRows, isComparing]);

    /** Saves the rows currently shown in the grid (search and filters applied). */
    const handleExport = useCallback(
        async (format: ExportFormat) => {
            try {
                const content = await buildExport(filteredRows, snapshots, format);
                const filter = EXPORT_FORMATS.find((f) => f.format === format)!.filter;
                const path = await window.toolboxAPI.fileSystem.saveFile(exportFileName(snapshots, format), content, [filter]);
                if (path) notify('Export complete', `Exported ${filteredRows.length} variables to ${path}`, 'success');
            } catch (error) {
                console.error('Error exporting environment variables:', error);
                notify('Export failed', describeError(error), 'error');
            }
        },
        [filteredRows, snapshots]
    );

    // Solutions that contain variables first (most first), then the rest alphabetically
    const solutionOptions = useMemo(() => {
        const counts = new Map<string, number>();
        rows.forEach((row) => row.solutions.forEach((name) => counts.set(name, (counts.get(name) ?? 0) + 1)));
        return solutions
            .map((solution) => ({ ...solution, count: counts.get(solution.uniqueName) ?? 0 }))
            .sort((a, b) => (b.count > 0 ? 1 : 0) - (a.count > 0 ? 1 : 0) || a.friendlyName.localeCompare(b.friendlyName));
    }, [rows, solutions]);

    const hasConnections = connections.some((c) => !!c);
    const fluentTheme = theme === 'dark' ? webDarkTheme : webLightTheme;

    if (isLoadingConnections || (hasConnections && !hasLoaded)) {
        const names = connections.filter((c): c is ToolBoxAPI.Connection => !!c).map((c) => c.name);
        return (
            <FluentProvider theme={fluentTheme}>
                <div className={styles.loadingScreen} role="status" aria-live="polite">
                    <Spinner size="small" />
                    <Text className={styles.loadingDetail}>
                        {isLoadingConnections ? 'Loading connections…' : `Loading environment variables from ${names.join(', ')}…`}
                    </Text>
                </div>
            </FluentProvider>
        );
    }

    let body: ReactNode;
    if (!hasConnections) {
        body = (
            <EmptyState
                icon={<PlugDisconnected24Regular />}
                title="No connection"
                description="Connect to a Dataverse environment to view its environment variables."
            />
        );
    } else if (rows.length === 0) {
        body = (
            <EmptyState
                icon={<DocumentBulletList24Regular />}
                title="No environment variables found"
                description="The connected environments don't contain any environment variable definitions."
            />
        );
    } else {
        body = (
            <VariablesTable
                rows={filteredRows}
                snapshots={snapshots}
                onSelect={selectRow}
                onCopy={copySingle}
                columnFilters={filters.columnFilters}
                onColumnFilterChange={setColumnFilter}
                // Keep the header (and its column filters) on screen so a filter that matches nothing can be edited
                emptyState={<EmptyState icon={<Search24Regular />} title="No matches" description="No environment variables match the current search and filters." />}
            />
        );
    }

    return (
        // Layout styles live on an inner div: Fluent copies the provider's className onto
        // portal mount nodes, so a full-height background there would cover the page
        // whenever a dropdown, tooltip or dialog opens.
        <FluentProvider theme={fluentTheme}>
            <div className={styles.root}>
                <header className={styles.topBar}>
                    <span className={styles.title}>
                        <span className={styles.titleIcon} aria-hidden>
                            <BracesVariable20Regular />
                        </span>
                        Environment Variables
                    </span>

                    {isLoading && <Spinner size="extra-tiny" />}
                    <div className={styles.environments}>
                        {snapshots.map((snapshot) => (
                            <span
                                key={snapshot.slot}
                                className={mergeClasses(styles.environment, snapshot.error && styles.environmentError)}
                                title={snapshot.connection.url}
                            >
                                <EnvDot color={snapshot.connection.environmentColor} />
                                <span className={styles.environmentName}>{snapshot.connection.name}</span>
                                {snapshot.connection.environment}
                            </span>
                        ))}
                    </div>
                    {hasConnections && (
                        <Button size="small" appearance="primary" icon={<Add16Regular />} onClick={() => setIsCreating(true)} disabled={isLoading}>
                            Create variable
                        </Button>
                    )}
                    {missingRows.length > 0 && (
                        <Tooltip content="Copy the missing variables shown to the environments that don't have them" relationship="description">
                            <Button size="small" appearance="subtle" icon={<CopyArrowRight16Regular />} onClick={() => openCopy(missingRows)} disabled={isLoading}>
                                Copy missing ({missingRows.length})
                            </Button>
                        </Tooltip>
                    )}
                    {hasConnections && (
                        <Menu>
                            <MenuTrigger disableButtonEnhancement>
                                <Button size="small" appearance="subtle" icon={<ArrowDownload16Regular />} disabled={isLoading || filteredRows.length === 0}>
                                    Export
                                </Button>
                            </MenuTrigger>
                            <MenuPopover>
                                <MenuList>
                                    {EXPORT_FORMATS.map(({ format, label }) => (
                                        <MenuItem key={format} onClick={() => handleExport(format)}>
                                            {label}
                                        </MenuItem>
                                    ))}
                                </MenuList>
                            </MenuPopover>
                        </Menu>
                    )}
                    {hasConnections && (
                        <Tooltip content="Reload from all environments" relationship="label">
                            <Button size="small" appearance="subtle" icon={<ArrowClockwise16Regular />} onClick={loadAll} disabled={isLoading} />
                        </Tooltip>
                    )}
                </header>

                {failed.length > 0 && (
                    <div className={styles.messages}>
                        {failed.map((snapshot) => (
                            <MessageBar key={snapshot.slot} intent="error">
                                <MessageBarBody>
                                    <MessageBarTitle>{snapshot.connection.name}</MessageBarTitle>
                                    {snapshot.error}
                                </MessageBarBody>
                            </MessageBar>
                        ))}
                    </div>
                )}

                <FilterBar
                    filters={filters}
                    onChange={setFilters}
                    solutions={solutionOptions}
                    isLoadingSolutions={isLoadingSolutions}
                    solutionsError={solutionsError}
                    isComparing={isComparing}
                />

                {body}

                {rows.length > 0 && (
                    <footer className={styles.footer}>
                        <Caption1 aria-live="polite">
                            {filteredRows.length === rows.length ? `${rows.length} variables` : `${filteredRows.length} of ${rows.length} variables`}
                        </Caption1>
                        {snapshots.length < MAX_ENVIRONMENTS && (
                            // Sits right above PPTB's numbered connection squares, which is where environments are added
                            <Popover positioning="above-end" withArrow>
                                <PopoverTrigger disableButtonEnhancement>
                                    <Button size="small" appearance="subtle" icon={<Add16Regular />} className={styles.addHint}>
                                        {isComparing ? 'Compare another environment' : 'Compare with another environment'}
                                        <span className={styles.addHintDetail}>
                                            · pick an empty numbered slot below <ArrowDown12Regular />
                                        </span>
                                    </Button>
                                </PopoverTrigger>
                                <PopoverSurface className={styles.addSurface}>
                                    <AddEnvironmentHelp used={snapshots.length} />
                                </PopoverSurface>
                            </Popover>
                        )}
                    </footer>
                )}

                <NewVariableDialog open={isCreating} snapshots={snapshots} onClose={() => setIsCreating(false)} onCreated={handleCreated} />

                <CopyVariablesDialog
                    rows={copyRequest ? rows.filter((row) => copyRequest.keys.includes(row.key)) : null}
                    snapshots={snapshots}
                    initialTarget={copyRequest?.target}
                    onClose={() => setCopyRequest(null)}
                    onCopied={handleCopied}
                />

                <VariableDetailsDialog
                    row={rows.find((row) => row.key === selectedKey) ?? null}
                    snapshots={snapshots}
                    initialEdit={editTarget}
                    onChanged={refreshVariable}
                    onCopy={copySingle}
                    onClose={() => setSelectedKey(null)}
                />
            </div>
        </FluentProvider>
    );
}

export default App;
