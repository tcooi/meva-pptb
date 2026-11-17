import React from 'react';
import {
    Button,
    Checkbox,
    Dialog,
    DialogActions,
    DialogBody,
    DialogContent,
    DialogSurface,
    DialogTitle,
    Field,
    MessageBar,
    MessageBarActions,
    MessageBarBody,
    Select,
    Spinner,
    Text,
    makeStyles,
    tokens,
} from '@fluentui/react-components';
import { Dismiss20Regular } from '@fluentui/react-icons';
import { notify } from '../hooks/useToolboxAPI';
import { useUnmanagedSolutions } from '../hooks/useUnmanagedSolutions';
import { canCopy, copyEnvironmentVariable, describeError, type EnvironmentSnapshot, type SolutionInfo, type VariableRow } from '../services/environmentVariables';
import { SolutionField } from './SolutionField';
import { TypePill } from './ui';

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
        alignItems: 'start',
    },
    list: {
        display: 'flex',
        flexDirection: 'column',
        maxHeight: '200px',
        overflowY: 'auto',
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        borderRadius: tokens.borderRadiusMedium,
    },
    item: {
        display: 'flex',
        alignItems: 'center',
        gap: tokens.spacingHorizontalS,
        minHeight: '30px',
        padding: `2px ${tokens.spacingHorizontalS}`,
        borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
        ':last-child': {
            borderBottom: 'none',
        },
    },
    itemName: {
        flex: 1,
        minWidth: 0,
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap',
    },
    itemSchema: {
        color: tokens.colorNeutralForeground3,
        fontSize: tokens.fontSizeBase200,
        marginLeft: tokens.spacingHorizontalXS,
    },
    note: {
        color: tokens.colorNeutralForeground3,
    },
});

interface CopyVariablesDialogProps {
    /** Rows to copy: one for a single variable, or the filtered missing rows for a bulk copy. Null when closed. */
    rows: VariableRow[] | null;
    snapshots: EnvironmentSnapshot[];
    /** Environment (connection slot) to copy to, when opened from a specific environment. */
    initialTarget?: number;
    onClose: () => void;
    /** Called once variables exist in the target, so the grid can show them. */
    onCopied: (slot: number, definitionIds: string[], solution: SolutionInfo | null) => Promise<void>;
}

const indexOfSlot = (snapshots: EnvironmentSnapshot[], slot: number) => snapshots.findIndex((s) => s.slot === slot);

export const CopyVariablesDialog: React.FC<CopyVariablesDialogProps> = ({ rows, snapshots, initialTarget, onClose, onCopied }) => {
    const styles = useStyles();
    const open = rows !== null;
    const list = rows ?? [];
    const loaded = snapshots.filter((s) => !s.error);
    const existsIn = (row: VariableRow, slot: number) => !!row.values[indexOfSlot(snapshots, slot)];

    // Targets: environments missing at least one of the rows. Sources: other environments holding at least one.
    const targets = loaded.filter((s) => list.some((row) => !existsIn(row, s.slot)));
    const [targetSlot, setTargetSlot] = React.useState<number | undefined>();
    const target = targets.find((s) => s.slot === targetSlot) ?? targets[0];
    const sources = target ? loaded.filter((s) => s.slot !== target.slot && list.some((row) => existsIn(row, s.slot) && !existsIn(row, target.slot))) : [];
    const [sourceSlot, setSourceSlot] = React.useState<number | undefined>();
    const source = sources.find((s) => s.slot === sourceSlot) ?? sources[0];

    const [includeCurrent, setIncludeCurrent] = React.useState(false);
    const [confirmingProduction, setConfirmingProduction] = React.useState(false);
    const [progress, setProgress] = React.useState<{ done: number; total: number } | null>(null);
    const [error, setError] = React.useState<string>();
    const busy = progress !== null;

    React.useEffect(() => {
        if (!open) return;
        setTargetSlot(initialTarget);
        setSourceSlot(undefined);
        setIncludeCurrent(false);
        setConfirmingProduction(false);
        setError(undefined);
    }, [open, initialTarget]);

    const candidates = target && source ? list.filter((row) => existsIn(row, source.slot) && !existsIn(row, target.slot)) : [];
    const toCopy = candidates.filter((row) => canCopy(row.type));
    const skipped = candidates.length - toCopy.length;

    // Prefer a solution with the same name as one the source variables are in, so ALM structure carries over
    const preferred = [...new Set(toCopy.flatMap((row) => row.solutions))];
    const solutionState = useUnmanagedSolutions(open, target?.slot, preferred);
    const { loading: loadingSolutions, solution } = solutionState;

    const canStart = !busy && !!target && !!source && toCopy.length > 0 && !loadingSolutions;

    const copy = async () => {
        if (!target || !source) return;
        if (target.connection.environment === 'Production' && !confirmingProduction) {
            setConfirmingProduction(true);
            return;
        }
        setError(undefined);
        const sourceIndex = indexOfSlot(snapshots, source.slot);
        const created: string[] = [];
        const problems: string[] = [];
        for (const [i, row] of toCopy.entries()) {
            setProgress({ done: i, total: toCopy.length });
            try {
                const { definitionId, warnings } = await copyEnvironmentVariable(target.slot, row.values[sourceIndex]!, {
                    includeCurrentValue: includeCurrent,
                    solutionUniqueName: solution?.uniqueName ?? null,
                });
                created.push(definitionId);
                warnings.forEach((w) => problems.push(`${row.schemaName}: ${w}`));
            } catch (err) {
                console.error(`Error copying ${row.schemaName} to ${target.connection.name}:`, err);
                problems.push(`${row.schemaName}: ${describeError(err)}`);
            }
        }

        try {
            if (created.length > 0) await onCopied(target.slot, created, solution);
        } catch (err) {
            console.error('Error refreshing copied variables:', err);
        }
        setProgress(null);
        setConfirmingProduction(false);

        if (created.length === 0) {
            // Nothing was written: stay open so the user can see why and try again
            setError(problems.join('\n'));
            return;
        }
        // Close first: the copied rows are no longer missing, so the dialog would otherwise empty out while the notification shows
        onClose();
        const summary = `${created.length} ${created.length === 1 ? 'variable' : 'variables'} copied to ${target.connection.name}`;
        if (problems.length > 0) {
            await notify('Copied with problems', `${summary}, but:\n${problems.join('\n')}`, 'warning');
        } else {
            await notify('Variables copied', `${summary}.`, 'success');
        }
    };

    const single = list.length === 1 ? list[0] : null;

    return (
        <Dialog open={open} onOpenChange={(_e, data) => !data.open && !busy && onClose()}>
            <DialogSurface className={styles.surface}>
                <DialogBody>
                    <DialogTitle action={<Button appearance="subtle" aria-label="Close" icon={<Dismiss20Regular />} onClick={onClose} disabled={busy} />}>
                        {single ? `Copy ${single.displayName}` : 'Copy missing variables'}
                    </DialogTitle>
                    <DialogContent className={styles.content}>
                        {targets.length === 0 ? (
                            <Text>These variables already exist in every environment.</Text>
                        ) : (
                            <>
                                <div className={styles.row}>
                                    <Field label="From">
                                        <Select value={String(source?.slot ?? '')} onChange={(_e, data) => setSourceSlot(Number(data.value))} disabled={busy || sources.length < 2}>
                                            {sources.map((s) => (
                                                <option key={s.slot} value={s.slot}>
                                                    {s.connection.name} ({s.connection.environment})
                                                </option>
                                            ))}
                                        </Select>
                                    </Field>
                                    <Field label="To">
                                        <Select value={String(target?.slot ?? '')} onChange={(_e, data) => setTargetSlot(Number(data.value))} disabled={busy || targets.length < 2}>
                                            {targets.map((s) => (
                                                <option key={s.slot} value={s.slot}>
                                                    {s.connection.name} ({s.connection.environment})
                                                </option>
                                            ))}
                                        </Select>
                                    </Field>
                                </div>

                                <Field label={`${toCopy.length} ${toCopy.length === 1 ? 'variable' : 'variables'} to copy`}>
                                    {toCopy.length > 0 && (
                                        <div className={styles.list}>
                                            {toCopy.map((row) => (
                                                <div key={row.key} className={styles.item}>
                                                    <span className={styles.itemName} title={row.schemaName}>
                                                        {row.displayName}
                                                        <span className={styles.itemSchema}>{row.schemaName}</span>
                                                    </span>
                                                    <TypePill type={row.type} />
                                                </div>
                                            ))}
                                        </div>
                                    )}
                                </Field>
                                {skipped > 0 && (
                                    <MessageBar intent="info">
                                        <MessageBarBody>
                                            {skipped} Secret or Data source {skipped === 1 ? 'variable is' : 'variables are'} skipped; create {skipped === 1 ? 'it' : 'them'} in the Power Apps maker
                                            portal.
                                        </MessageBarBody>
                                    </MessageBar>
                                )}

                                <SolutionField state={solutionState} hint="Schema names keep their original prefix." disabled={busy} />

                                <Field hint="Current values are usually specific to each environment, so they're left out unless you ask.">
                                    <Checkbox checked={includeCurrent} onChange={(_e, data) => setIncludeCurrent(!!data.checked)} label="Also copy current values" disabled={busy} />
                                </Field>

                                {confirmingProduction && target && (
                                    <MessageBar intent="warning">
                                        <MessageBarBody>{target.connection.name} is a Production environment. Copy these variables there?</MessageBarBody>
                                        <MessageBarActions>
                                            <Button size="small" appearance="primary" onClick={copy} disabled={busy}>
                                                Copy to Production
                                            </Button>
                                            <Button size="small" onClick={() => setConfirmingProduction(false)} disabled={busy}>
                                                Back
                                            </Button>
                                        </MessageBarActions>
                                    </MessageBar>
                                )}
                                {error && (
                                    <MessageBar intent="error">
                                        <MessageBarBody style={{ whiteSpace: 'pre-line' }}>{error}</MessageBarBody>
                                    </MessageBar>
                                )}
                            </>
                        )}
                    </DialogContent>
                    <DialogActions>
                        {progress && (
                            <>
                                <Spinner size="tiny" />
                                <Text className={styles.note} aria-live="polite">
                                    Copying {progress.done + 1} of {progress.total}…
                                </Text>
                            </>
                        )}
                        <Button onClick={onClose} disabled={busy}>
                            Cancel
                        </Button>
                        <Button appearance="primary" onClick={copy} disabled={!canStart}>
                            Copy
                        </Button>
                    </DialogActions>
                </DialogBody>
            </DialogSurface>
        </Dialog>
    );
};
