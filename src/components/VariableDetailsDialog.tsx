import React from 'react';
import {
    Button,
    Dialog,
    DialogActions,
    DialogBody,
    DialogContent,
    DialogSurface,
    DialogTitle,
    MessageBar,
    MessageBarActions,
    MessageBarBody,
    Select,
    Spinner,
    Text,
    Textarea,
    Tooltip,
    makeStyles,
    mergeClasses,
    tokens,
} from '@fluentui/react-components';
import { Box16Regular, BranchCompare16Regular, CopyArrowRight16Regular, Checkmark12Filled, Checkmark16Regular, Copy16Regular, Dismiss20Regular, Edit16Regular, ErrorCircle16Regular, Shapes16Regular, TextDescription16Regular, TextT16Regular } from '@fluentui/react-icons';
import {
    BOOLEAN_VALUES,
    canCopy,
    describeError,
    removeCurrentValue,
    saveCurrentValue,
    saveDefaultValue,
    validateValue,
    type EditTarget,
    type EnvironmentSnapshot,
    type EnvironmentVariable,
    type ValueKind,
    type VariableRow,
} from '../services/environmentVariables';
import { EnvDot, NO_TEXT_ASSIST, StatusPill, TypePill } from './ui';

const useStyles = makeStyles({
    surface: {
        maxWidth: '760px',
        width: '92vw',
    },
    titleText: {
        display: 'flex',
        flexDirection: 'column',
        minWidth: 0,
    },
    schemaName: {
        fontSize: tokens.fontSizeBase200,
        fontWeight: tokens.fontWeightRegular,
        color: tokens.colorNeutralForeground3,
    },
    content: {
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalL,
        maxHeight: '65vh',
        overflowY: 'auto',
    },
    field: {
        display: 'flex',
        flexDirection: 'column',
        gap: '6px',
    },
    // Small grey field label with an icon, as in an expanded record
    label: {
        display: 'flex',
        alignItems: 'center',
        gap: '6px',
        fontSize: tokens.fontSizeBase200,
        color: tokens.colorNeutralForeground3,
    },
    inline: {
        display: 'flex',
        flexWrap: 'wrap',
        gap: tokens.spacingHorizontalXS,
        alignItems: 'center',
    },
    solution: {
        display: 'inline-flex',
        alignItems: 'center',
        height: '20px',
        padding: '0 8px',
        borderRadius: tokens.borderRadiusCircular,
        fontSize: tokens.fontSizeBase200,
        backgroundColor: tokens.colorNeutralBackground3,
        color: tokens.colorNeutralForeground2,
    },
    section: {
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalM,
        paddingTop: tokens.spacingVerticalL,
        borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    },
    sectionHeader: {
        display: 'flex',
        alignItems: 'center',
        gap: tokens.spacingHorizontalS,
        minWidth: 0,
    },
    sectionName: {
        fontWeight: tokens.fontWeightSemibold,
        flexShrink: 0,
    },
    sectionMeta: {
        color: tokens.colorNeutralForeground3,
        fontSize: tokens.fontSizeBase200,
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap',
        minWidth: 0,
    },
    valueBox: {
        display: 'flex',
        alignItems: 'flex-start',
        gap: tokens.spacingHorizontalS,
        minHeight: '32px',
        padding: `6px ${tokens.spacingHorizontalS}`,
        borderRadius: tokens.borderRadiusMedium,
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        backgroundColor: tokens.colorNeutralBackground1,
    },
    value: {
        flex: 1,
        margin: 0,
        paddingTop: '2px',
        fontFamily: tokens.fontFamilyBase,
        fontSize: tokens.fontSizeBase300,
        lineHeight: tokens.lineHeightBase300,
        whiteSpace: 'pre-wrap',
        wordBreak: 'break-all',
        maxHeight: '180px',
        overflowY: 'auto',
    },
    json: {
        fontFamily: tokens.fontFamilyMonospace,
        fontSize: tokens.fontSizeBase200,
    },
    savedNote: {
        display: 'inline-flex',
        alignItems: 'center',
        gap: '2px',
        color: tokens.colorStatusSuccessForeground1,
        fontWeight: tokens.fontWeightSemibold,
    },
    inUse: {
        color: tokens.colorBrandForeground1,
        fontWeight: tokens.fontWeightSemibold,
    },
    editor: {
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalS,
    },
    editorInput: {
        width: '100%',
        '& textarea': {
            fontFamily: tokens.fontFamilyBase,
            maxHeight: '240px',
        },
    },
    editorInputJson: {
        '& textarea': {
            fontFamily: tokens.fontFamilyMonospace,
            fontSize: tokens.fontSizeBase200,
        },
    },
    editorActions: {
        display: 'flex',
        alignItems: 'center',
        gap: tokens.spacingHorizontalS,
    },
    editorSpacer: {
        flex: 1,
    },
    error: {
        color: tokens.colorStatusDangerForeground1,
    },
    note: {
        color: tokens.colorNeutralForeground3,
    },
    empty: {
        color: tokens.colorNeutralForeground4,
        paddingTop: '2px',
    },
    missing: {
        display: 'flex',
        flexWrap: 'wrap',
        alignItems: 'center',
        justifyContent: 'space-between',
        gap: tokens.spacingHorizontalS,
    },
});

function format(value: string, type: EnvironmentVariable['type']): string {
    if (type !== 'JSON') return value;
    try {
        return JSON.stringify(JSON.parse(value), null, 2);
    } catch {
        return value;
    }
}

type CopyState = 'idle' | 'copied' | 'failed';

/**
 * Copies text and reports the result in place (for about 1.5 s) instead of as a PPTB notification,
 * which would pile up in PPTB's notification list for something this routine.
 */
function useCopy() {
    const [state, setState] = React.useState<CopyState>('idle');
    const timer = React.useRef<ReturnType<typeof setTimeout>>();
    React.useEffect(() => () => clearTimeout(timer.current), []);

    const copy = React.useCallback(async (text: string) => {
        let next: CopyState = 'copied';
        try {
            await window.toolboxAPI.utils.copyToClipboard(text);
        } catch (error) {
            console.error('Error copying to clipboard:', error);
            next = 'failed';
        }
        setState(next);
        clearTimeout(timer.current);
        timer.current = setTimeout(() => setState('idle'), 1500);
    }, []);

    return { state, copy };
}

const COPY_ICONS: Record<CopyState, React.ReactElement> = {
    idle: <Copy16Regular />,
    copied: <Checkmark16Regular />,
    failed: <ErrorCircle16Regular />,
};

/** Icon button that turns into a check mark (or an error icon) after copying. */
const CopyButton: React.FC<{ text: string; what: string }> = ({ text, what }) => {
    const { state, copy } = useCopy();
    const tooltip = state === 'copied' ? 'Copied' : state === 'failed' ? "Couldn't copy" : `Copy ${what}`;
    return (
        <Tooltip content={tooltip} relationship="label" visible={state !== 'idle' ? true : undefined}>
            <Button size="small" appearance="subtle" icon={COPY_ICONS[state]} onClick={() => copy(text)} />
        </Tooltip>
    );
};

const Field: React.FC<{ icon: React.ReactElement; label: React.ReactNode; children: React.ReactNode }> = ({ icon, label, children }) => {
    const styles = useStyles();
    return (
        <div className={styles.field}>
            <span className={styles.label}>
                {icon}
                {label}
            </span>
            {children}
        </div>
    );
};

interface ValueFieldProps {
    kind: ValueKind;
    variable: EnvironmentVariable;
    snapshot: EnvironmentSnapshot;
    onChanged: (slot: number, definitionId: string) => Promise<unknown>;
    /** Open in edit mode straight away (the dialog was opened from a grid cell's edit button). */
    autoEdit?: boolean;
}

/** Shows a current/default value with copy and edit; switches to an inline editor when editing. */
const ValueField: React.FC<ValueFieldProps> = ({ kind, variable, snapshot, onChanged, autoEdit }) => {
    const styles = useStyles();
    const value = kind === 'current' ? variable.currentValue : variable.defaultValue;
    const label = kind === 'current' ? 'Current value' : 'Default value';
    const inUse = kind === 'current' ? variable.currentValue !== null : variable.currentValue === null && variable.defaultValue !== null;
    const isProduction = snapshot.connection.environment === 'Production';

    const [editing, setEditing] = React.useState(false);
    const [draft, setDraft] = React.useState('');
    // An action waiting for "Save to Production" confirmation
    const [pending, setPending] = React.useState<'save' | 'remove' | null>(null);
    const [busy, setBusy] = React.useState(false);
    const [error, setError] = React.useState<string>();
    // Brief "Saved"/"Removed" note next to the label instead of a PPTB notification
    const [savedNote, setSavedNote] = React.useState<string | null>(null);
    const savedTimer = React.useRef<ReturnType<typeof setTimeout>>();
    React.useEffect(() => () => clearTimeout(savedTimer.current), []);
    // The editor's buttons disappear when it closes; move focus back to Edit so it stays inside the
    // dialog (otherwise Escape and Tab stop working until the user clicks into the dialog again)
    const editButtonRef = React.useRef<HTMLButtonElement>(null);
    const wasEditing = React.useRef(false);
    React.useEffect(() => {
        if (wasEditing.current && !editing) editButtonRef.current?.focus();
        wasEditing.current = editing;
    }, [editing]);

    const startEditing = () => {
        setDraft(value ?? (variable.type === 'Boolean' ? 'yes' : ''));
        setPending(null);
        setError(undefined);
        setEditing(true);
    };

    // Opened from a grid cell's edit button: start editing and bring the editor into view
    const fieldRef = React.useRef<HTMLDivElement>(null);
    React.useEffect(() => {
        if (autoEdit && variable.type !== 'Secret') {
            startEditing();
            requestAnimationFrame(() => fieldRef.current?.scrollIntoView({ block: 'center' }));
        }
        // Only on mount: the field is keyed per variable, so a new dialog mounts a new field
    }, []);

    const validation = draft === '' ? undefined : validateValue(variable.type, variable.type === 'Number' ? draft.trim() : draft);
    const canSave = !busy && draft !== '' && !validation && draft !== value;
    const canRemove = kind === 'current' ? variable.currentValueId !== null : variable.defaultValue !== null;

    const run = async (action: 'save' | 'remove') => {
        if (isProduction && pending !== action) {
            setPending(action);
            return;
        }
        setBusy(true);
        setError(undefined);
        try {
            const target = snapshot.slot;
            const newValue = variable.type === 'Number' ? draft.trim() : draft;
            if (kind === 'current') {
                await (action === 'save' ? saveCurrentValue(target, variable, newValue) : removeCurrentValue(target, variable));
            } else {
                await saveDefaultValue(target, variable, action === 'save' ? newValue : null);
            }
            await onChanged(snapshot.slot, variable.definitionId);
            setSavedNote(action === 'save' ? 'Saved' : kind === 'current' ? 'Removed' : 'Cleared');
            clearTimeout(savedTimer.current);
            savedTimer.current = setTimeout(() => setSavedNote(null), 2500);
            setEditing(false);
        } catch (err) {
            console.error(`Error changing ${label.toLowerCase()} of ${variable.schemaName} in ${snapshot.connection.name}:`, err);
            setError(describeError(err));
        } finally {
            setBusy(false);
            setPending(null);
        }
    };

    const header = (
        <>
            {label}
            {inUse && <span className={styles.inUse}>· In use</span>}
            <span className={styles.savedNote} aria-live="polite">
                {savedNote && (
                    <>
                        <Checkmark12Filled />
                        {savedNote}
                    </>
                )}
            </span>
        </>
    );

    if (!editing) {
        return (
            <Field icon={<TextT16Regular />} label={header}>
                <div className={styles.valueBox}>
                    {value === null ? (
                        <span className={styles.empty}>Not set</span>
                    ) : (
                        <pre className={mergeClasses(styles.value, variable.type === 'JSON' && styles.json)}>{value === '' ? '(empty)' : format(value, variable.type)}</pre>
                    )}
                    {value === null && <span className={styles.editorSpacer} />}
                    {value !== null && <CopyButton text={value} what={label.toLowerCase()} />}
                    {variable.type !== 'Secret' && (
                        <Tooltip content={`Edit ${label.toLowerCase()}`} relationship="label">
                            <Button ref={editButtonRef} size="small" appearance="subtle" icon={<Edit16Regular />} onClick={startEditing} />
                        </Tooltip>
                    )}
                </div>
                {variable.type === 'Secret' && (
                    <Text size={200} className={styles.note}>
                        Secret values can't be edited here; change them in the Power Apps maker portal.
                    </Text>
                )}
            </Field>
        );
    }

    return (
        <Field icon={<TextT16Regular />} label={header}>
            <div className={styles.editor} ref={fieldRef}>
                {kind === 'default' && variable.isManaged && (
                    <MessageBar intent="warning">
                        <MessageBarBody>
                            This variable comes from a managed solution. Changing its default here adds an unmanaged customization on top of the solution, or
                            may be refused. To change the value for this environment, set a current value instead.
                        </MessageBarBody>
                    </MessageBar>
                )}

                {variable.type === 'Boolean' ? (
                    <Select aria-label={label} value={draft} onChange={(_e, data) => setDraft(data.value)} disabled={busy} autoFocus>
                        {BOOLEAN_VALUES.map((option) => (
                            <option key={option} value={option}>
                                {option === 'yes' ? 'Yes' : 'No'}
                            </option>
                        ))}
                    </Select>
                ) : (
                    <Textarea
                        {...NO_TEXT_ASSIST}
                        aria-label={label}
                        className={mergeClasses(styles.editorInput, variable.type === 'JSON' && styles.editorInputJson)}
                        resize="vertical"
                        rows={variable.type === 'JSON' ? 6 : 2}
                        autoFocus
                        value={draft}
                        disabled={busy}
                        onChange={(_e, data) => setDraft(data.value)}
                    />
                )}
                {validation && (
                    <Text size={200} className={styles.error}>
                        {validation}
                    </Text>
                )}
                {kind === 'current' && draft === '' && (
                    <Text size={200} className={styles.note}>
                        To fall back to the default value, use Remove current value.
                    </Text>
                )}

                {pending && (
                    <MessageBar intent="warning">
                        <MessageBarBody>
                            {snapshot.connection.name} is a Production environment.{' '}
                            {pending === 'save' ? `Save this ${label.toLowerCase()}?` : kind === 'current' ? 'Remove the current value?' : 'Clear the default value?'}
                        </MessageBarBody>
                        <MessageBarActions>
                            <Button size="small" appearance="primary" onClick={() => run(pending)} disabled={busy}>
                                {pending === 'save' ? 'Save to Production' : 'Yes, in Production'}
                            </Button>
                            <Button size="small" onClick={() => setPending(null)} disabled={busy}>
                                Back
                            </Button>
                        </MessageBarActions>
                    </MessageBar>
                )}
                {error && (
                    <MessageBar intent="error">
                        <MessageBarBody>{error}</MessageBarBody>
                    </MessageBar>
                )}

                <div className={styles.editorActions}>
                    {canRemove && (
                        <Button size="small" appearance="subtle" disabled={busy} onClick={() => run('remove')}>
                            {kind === 'current' ? 'Remove current value' : 'Clear default'}
                        </Button>
                    )}
                    <span className={styles.editorSpacer} />
                    {busy && <Spinner size="extra-tiny" />}
                    <Button size="small" disabled={busy} onClick={() => setEditing(false)}>
                        Cancel
                    </Button>
                    <Button size="small" appearance="primary" disabled={!canSave} onClick={() => run('save')}>
                        Save
                    </Button>
                </div>
            </div>
        </Field>
    );
};

interface VariableDetailsDialogProps {
    row: VariableRow | null;
    snapshots: EnvironmentSnapshot[];
    /** Value to open straight in edit mode, when the dialog was opened from a grid cell's edit button. */
    initialEdit?: EditTarget | null;
    /** Called after a value is saved so the variable can be re-read from that environment. */
    onChanged: (slot: number, definitionId: string) => Promise<unknown>;
    /** Opens the copy dialog to create this variable in an environment that doesn't have it. */
    onCopy: (row: VariableRow, slot: number) => void;
    onClose: () => void;
}

export const VariableDetailsDialog: React.FC<VariableDetailsDialogProps> = ({ row, snapshots, initialEdit, onChanged, onCopy, onClose }) => {
    const styles = useStyles();
    const schemaCopy = useCopy();
    if (!row) return null;

    return (
        <Dialog open onOpenChange={(_e, data) => !data.open && onClose()}>
            <DialogSurface className={styles.surface}>
                <DialogBody>
                    <DialogTitle action={<Button appearance="subtle" aria-label="Close" icon={<Dismiss20Regular />} onClick={onClose} />}>
                        <span className={styles.titleText}>
                            {row.displayName}
                            <span className={styles.schemaName}>{row.schemaName}</span>
                        </span>
                    </DialogTitle>
                    <DialogContent className={styles.content}>
                        <Field icon={<Shapes16Regular />} label="Type">
                            <span className={styles.inline}>
                                <TypePill type={row.type} />
                            </span>
                        </Field>
                        {row.status !== 'single' && (
                            <Field icon={<BranchCompare16Regular />} label="Comparison">
                                <span className={styles.inline}>
                                    <StatusPill status={row.status} />
                                </span>
                            </Field>
                        )}
                        {row.solutions.length > 0 && (
                            <Field icon={<Box16Regular />} label="Solutions">
                                <span className={styles.inline}>
                                    {row.solutions.map((solution) => (
                                        <span key={solution} className={styles.solution}>
                                            {solution}
                                        </span>
                                    ))}
                                </span>
                            </Field>
                        )}
                        {row.description && (
                            <Field icon={<TextDescription16Regular />} label="Description">
                                <Text>{row.description}</Text>
                            </Field>
                        )}

                        {snapshots.map((snapshot, index) => {
                            const variable = row.values[index];
                            return (
                                <section key={snapshot.slot} className={styles.section}>
                                    <div className={styles.sectionHeader}>
                                        <EnvDot color={snapshot.connection.environmentColor} />
                                        <Text className={styles.sectionName}>{snapshot.connection.name}</Text>
                                        <span className={styles.sectionMeta} title={snapshot.connection.url}>
                                            {variable ? `${variable.isManaged ? 'Managed' : 'Unmanaged'} · ` : ''}
                                            {snapshot.connection.url}
                                        </span>
                                    </div>
                                    {snapshot.error ? (
                                        <Text className={styles.empty}>Failed to load: {snapshot.error}</Text>
                                    ) : variable ? (
                                        <>
                                            <ValueField key={`${variable.definitionId}-current`} kind="current" variable={variable} snapshot={snapshot} onChanged={onChanged} autoEdit={initialEdit?.slot === snapshot.slot && initialEdit.kind === 'current'} />
                                            <ValueField key={`${variable.definitionId}-default`} kind="default" variable={variable} snapshot={snapshot} onChanged={onChanged} autoEdit={initialEdit?.slot === snapshot.slot && initialEdit.kind === 'default'} />
                                        </>
                                    ) : (
                                        <div className={styles.missing}>
                                            <Text className={styles.empty}>This variable doesn't exist in this environment.</Text>
                                            {canCopy(row.type) ? (
                                                <Button size="small" icon={<CopyArrowRight16Regular />} onClick={() => onCopy(row, snapshot.slot)}>
                                                    Copy to {snapshot.connection.name}
                                                </Button>
                                            ) : (
                                                <Text size={200} className={styles.note}>
                                                    {row.type} variables can't be copied here; create it in the Power Apps maker portal.
                                                </Text>
                                            )}
                                        </div>
                                    )}
                                </section>
                            );
                        })}
                    </DialogContent>
                    <DialogActions>
                        <Button size="small" icon={COPY_ICONS[schemaCopy.state]} onClick={() => schemaCopy.copy(row.schemaName)}>
                            {schemaCopy.state === 'copied' ? 'Copied' : schemaCopy.state === 'failed' ? "Couldn't copy" : 'Copy schema name'}
                        </Button>
                        <Button size="small" appearance="primary" onClick={onClose}>
                            Close
                        </Button>
                    </DialogActions>
                </DialogBody>
            </DialogSurface>
        </Dialog>
    );
};
