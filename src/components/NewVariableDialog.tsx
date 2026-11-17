import React from 'react';
import {
    Button,
    Dialog,
    DialogActions,
    DialogBody,
    DialogContent,
    DialogSurface,
    DialogTitle,
    Field,
    Input,
    MessageBar,
    MessageBarActions,
    MessageBarBody,
    Select,
    Spinner,
    Textarea,
    makeStyles,
    tokens,
} from '@fluentui/react-components';
import { Dismiss20Regular } from '@fluentui/react-icons';
import { notify } from '../hooks/useToolboxAPI';
import {
    BOOLEAN_VALUES,
    CREATABLE_TYPES,
    createEnvironmentVariable,
    describeError,
    suggestSchemaName,
    validateSchemaName,
    validateValue,
    type EnvironmentSnapshot,
    type UnmanagedSolution,
    type VariableType,
} from '../services/environmentVariables';
import { useUnmanagedSolutions } from '../hooks/useUnmanagedSolutions';
import { SolutionField } from './SolutionField';
import { NO_TEXT_ASSIST } from './ui';

const useStyles = makeStyles({
    surface: {
        maxWidth: '620px',
        width: '92vw',
    },
    content: {
        display: 'flex',
        flexDirection: 'column',
        gap: tokens.spacingVerticalM,
        maxHeight: '70vh',
        overflowY: 'auto',
    },
    row: {
        display: 'grid',
        gridTemplateColumns: '1fr 1fr',
        gap: tokens.spacingHorizontalM,
        // Keep side-by-side fields level when only one of them has a hint underneath
        alignItems: 'start',
    },
    prefix: {
        color: tokens.colorNeutralForeground3,
        fontFamily: tokens.fontFamilyBase,
        fontSize: tokens.fontSizeBase300,
    },
    // Fluent pads both the prefix slot and the input; drop that so "contoso_" and the name read as one word
    schemaInputRoot: {
        columnGap: 0,
        '& > .fui-Input__contentBefore': {
            paddingRight: 0,
            // Same inset as the text inside the other inputs, so "contoso_" starts where their text does
            paddingLeft: tokens.spacingHorizontalXXS,
        },
    },
    schemaInputField: {
        paddingLeft: 0,
    },
    json: {
        '& textarea': {
            fontFamily: tokens.fontFamilyMonospace,
            fontSize: tokens.fontSizeBase200,
        },
    },
});

interface ValueInputProps {
    label: string;
    hint: string;
    type: VariableType;
    value: string;
    onChange: (value: string) => void;
    disabled: boolean;
}

/** Optional value input that follows the variable type: Yes/No choice for Boolean, multi-line for JSON. */
const ValueInput: React.FC<ValueInputProps> = ({ label, hint, type, value, onChange, disabled }) => {
    const styles = useStyles();
    const validation = value === '' ? undefined : validateValue(type, type === 'Number' ? value.trim() : value);
    return (
        <Field label={label} hint={hint} validationMessage={validation} validationState={validation ? 'error' : 'none'}>
            {type === 'Boolean' ? (
                <Select value={value} onChange={(_e, data) => onChange(data.value)} disabled={disabled}>
                    <option value="">Not set</option>
                    {BOOLEAN_VALUES.map((option) => (
                        <option key={option} value={option}>
                            {option === 'yes' ? 'Yes' : 'No'}
                        </option>
                    ))}
                </Select>
            ) : type === 'JSON' ? (
                <Textarea
                    {...NO_TEXT_ASSIST}
                    className={styles.json}
                    resize="vertical"
                    rows={4}
                    value={value}
                    onChange={(_e, data) => onChange(data.value)}
                    disabled={disabled}
                />
            ) : (
                // One-line input for text and numbers, so it lines up with the other fields
                <Input {...NO_TEXT_ASSIST} value={value} onChange={(_e, data) => onChange(data.value)} disabled={disabled} />
            )}
        </Field>
    );
};

interface NewVariableDialogProps {
    open: boolean;
    snapshots: EnvironmentSnapshot[];
    onClose: () => void;
    /** Called once the definition exists, so the grid can show it and its details can open. */
    onCreated: (slot: number, definitionId: string, solution: UnmanagedSolution | null) => Promise<void>;
}

export const NewVariableDialog: React.FC<NewVariableDialogProps> = ({ open, snapshots, onClose, onCreated }) => {
    const styles = useStyles();
    const available = snapshots.filter((s) => !s.error);

    const [slot, setSlot] = React.useState<number>(available[0]?.slot ?? 0);
    const [manualPrefix, setManualPrefix] = React.useState('new');
    const [displayName, setDisplayName] = React.useState('');
    const [namePart, setNamePart] = React.useState('');
    // Until the user edits the schema name, it follows the display name
    const [nameEdited, setNameEdited] = React.useState(false);
    const [description, setDescription] = React.useState('');
    const [type, setType] = React.useState<VariableType>('String');
    const [defaultValue, setDefaultValue] = React.useState('');
    const [currentValue, setCurrentValue] = React.useState('');
    const [confirmingProduction, setConfirmingProduction] = React.useState(false);
    const [busy, setBusy] = React.useState(false);
    const [error, setError] = React.useState<string>();

    const snapshot = available.find((s) => s.slot === slot) ?? available[0];

    // Start fresh each time the dialog opens
    React.useEffect(() => {
        if (!open) return;
        setSlot(available[0]?.slot ?? 0);
        setDisplayName('');
        setNamePart('');
        setNameEdited(false);
        setDescription('');
        setType('String');
        setDefaultValue('');
        setCurrentValue('');
        setConfirmingProduction(false);
        setError(undefined);
    }, [open]);

    // Solutions (and their publisher prefixes) are per environment
    const solutionState = useUnmanagedSolutions(open, snapshot?.slot);
    const { solutions, loading: loadingSolutions, solution } = solutionState;
    // Before the first solution list arrives there's no prefix yet; show a placeholder rather than the "new" fallback
    const prefix = solution ? solution.prefix : loadingSolutions ? '…' : manualPrefix.trim();
    // "No solution" was actually picked (not just "nothing loaded yet"), so a prefix has to be typed in
    const needsManualPrefix = !loadingSolutions && solutions !== null && !solution;
    const schemaName = `${prefix}_${namePart}`;

    const schemaError = namePart ? validateSchemaName(schemaName, snapshot?.variables ?? new Map()) : undefined;
    const valueProblem = (value: string) => value !== '' && !!validateValue(type, type === 'Number' ? value.trim() : value);
    const canCreate =
        !busy && !!snapshot && !loadingSolutions && solutions !== null && displayName.trim() !== '' && namePart !== '' && !schemaError && !valueProblem(defaultValue) && !valueProblem(currentValue);

    const changeType = (next: VariableType) => {
        setType(next);
        // Values typed for one type rarely fit another
        setDefaultValue('');
        setCurrentValue('');
    };

    const create = async () => {
        if (!snapshot) return;
        if (snapshot.connection.environment === 'Production' && !confirmingProduction) {
            setConfirmingProduction(true);
            return;
        }
        setBusy(true);
        setError(undefined);
        const clean = (value: string) => (value === '' ? null : type === 'Number' ? value.trim() : value);
        try {
            const { definitionId, warnings } = await createEnvironmentVariable(snapshot.slot, {
                schemaName,
                displayName: displayName.trim(),
                description: description.trim(),
                type,
                defaultValue: clean(defaultValue),
                currentValue: clean(currentValue),
                solutionUniqueName: solution?.uniqueName ?? null,
            });
            await onCreated(snapshot.slot, definitionId, solution);
            // Success needs no notification: the new variable's details open. Only a half-finished create is reported.
            if (warnings.length > 0) {
                await notify('Variable created with problems', `${displayName.trim()} was created in ${snapshot.connection.name}, but:\n${warnings.join('\n')}`, 'warning');
            }
            onClose();
        } catch (err) {
            console.error('Error creating environment variable:', err);
            setError(describeError(err));
        } finally {
            setBusy(false);
            setConfirmingProduction(false);
        }
    };

    return (
        <Dialog open={open} onOpenChange={(_e, data) => !data.open && !busy && onClose()}>
            <DialogSurface className={styles.surface}>
                <DialogBody>
                    <DialogTitle action={<Button appearance="subtle" aria-label="Close" icon={<Dismiss20Regular />} onClick={onClose} disabled={busy} />}>
                        Create environment variable
                    </DialogTitle>
                    <DialogContent className={styles.content}>
                        <div className={styles.row}>
                            {available.length > 1 && (
                                <Field label="Environment">
                                    <Select value={String(slot)} onChange={(_e, data) => setSlot(Number(data.value))} disabled={busy}>
                                        {available.map((s) => (
                                            <option key={s.slot} value={s.slot}>
                                                {s.connection.name} ({s.connection.environment})
                                            </option>
                                        ))}
                                    </Select>
                                </Field>
                            )}
                            <SolutionField
                                state={solutionState}
                                hint={loadingSolutions ? 'Publisher prefix: …' : solution ? `Publisher prefix: ${solution.prefix}` : 'Enter a prefix below.'}
                                disabled={busy}
                            />
                        </div>

                        <Field label="Display name" required>
                            <Input
                                {...NO_TEXT_ASSIST}
                                autoFocus
                                value={displayName}
                                disabled={busy}
                                onChange={(_e, data) => {
                                    setDisplayName(data.value);
                                    if (!nameEdited) setNamePart(suggestSchemaName(data.value));
                                }}
                            />
                        </Field>

                        {needsManualPrefix && (
                            <Field label="Prefix" hint="No solution selected, so enter the publisher prefix to use.">
                                <Input {...NO_TEXT_ASSIST} value={manualPrefix} onChange={(_e, data) => setManualPrefix(data.value)} disabled={busy} />
                            </Field>
                        )}

                        <div className={styles.row}>
                            <Field
                                label="Schema name"
                                required
                                validationMessage={schemaError}
                                validationState={schemaError ? 'error' : 'none'}
                                hint={schemaError ? undefined : `Saved as ${schemaName}`}
                            >
                                <Input
                                    {...NO_TEXT_ASSIST}
                                    className={styles.schemaInputRoot}
                                    input={{ className: styles.schemaInputField }}
                                    contentBefore={<span className={styles.prefix}>{prefix}_</span>}
                                    value={namePart}
                                    disabled={busy}
                                    onChange={(_e, data) => {
                                        setNamePart(data.value);
                                        setNameEdited(true);
                                    }}
                                />
                            </Field>
                            <Field label="Type">
                                <Select value={type} onChange={(_e, data) => changeType(data.value as VariableType)} disabled={busy}>
                                    {CREATABLE_TYPES.map((t) => (
                                        <option key={t} value={t}>
                                            {t === 'Boolean' ? 'Yes/No' : t === 'String' ? 'Text' : t}
                                        </option>
                                    ))}
                                </Select>
                            </Field>
                        </div>

                        <Field label="Description">
                            <Textarea {...NO_TEXT_ASSIST} resize="vertical" rows={2} value={description} onChange={(_e, data) => setDescription(data.value)} disabled={busy} />
                        </Field>

                        <ValueInput
                            label="Default value"
                            hint="Travels with the solution to other environments."
                            type={type}
                            value={defaultValue}
                            onChange={setDefaultValue}
                            disabled={busy}
                        />
                        <ValueInput
                            label="Current value"
                            hint="Only for this environment; overrides the default here."
                            type={type}
                            value={currentValue}
                            onChange={setCurrentValue}
                            disabled={busy}
                        />

                        {confirmingProduction && snapshot && (
                            <MessageBar intent="warning">
                                <MessageBarBody>{snapshot.connection.name} is a Production environment. Create this variable there?</MessageBarBody>
                                <MessageBarActions>
                                    <Button size="small" appearance="primary" onClick={create} disabled={busy}>
                                        Create in Production
                                    </Button>
                                    <Button size="small" onClick={() => setConfirmingProduction(false)} disabled={busy}>
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
                    </DialogContent>
                    <DialogActions>
                        {busy && <Spinner size="tiny" />}
                        <Button onClick={onClose} disabled={busy}>
                            Cancel
                        </Button>
                        <Button appearance="primary" onClick={create} disabled={!canCreate}>
                            Create
                        </Button>
                    </DialogActions>
                </DialogBody>
            </DialogSurface>
        </Dialog>
    );
};
