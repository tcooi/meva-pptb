import React from 'react';
import { Field, Select, Spinner, makeStyles, tokens } from '@fluentui/react-components';
import { NO_SOLUTION, type useUnmanagedSolutions } from '../hooks/useUnmanagedSolutions';

const useStyles = makeStyles({
    labelWithSpinner: {
        display: 'inline-flex',
        alignItems: 'center',
        gap: tokens.spacingHorizontalXS,
    },
});

interface SolutionFieldProps {
    state: ReturnType<typeof useUnmanagedSolutions>;
    /** Hint under the picker once solutions have loaded. */
    hint: string;
    disabled: boolean;
}

/** Unmanaged solution picker. The dropdown and its hint line are always rendered, so a dialog keeps its size while solutions load. */
export const SolutionField: React.FC<SolutionFieldProps> = ({ state, hint, disabled }) => {
    const styles = useStyles();
    const { solutions, loading, error, selected, setSelected } = state;
    return (
        <Field
            label={
                <span className={styles.labelWithSpinner}>
                    Solution
                    {loading && <Spinner size="extra-tiny" aria-label="Loading solutions" />}
                </span>
            }
            hint={error ? undefined : hint}
            validationMessage={error ? `Couldn't load solutions: ${error}` : undefined}
            validationState={error ? 'warning' : 'none'}
        >
            <Select value={loading ? '__loading' : selected} onChange={(_e, data) => setSelected(data.value)} disabled={disabled || loading}>
                {loading ? (
                    <option value="__loading">Loading solutions…</option>
                ) : (
                    <>
                        {(solutions ?? []).map((s) => (
                            <option key={s.uniqueName} value={s.uniqueName}>
                                {s.friendlyName}
                            </option>
                        ))}
                        <option value={NO_SOLUTION}>No solution</option>
                    </>
                )}
            </Select>
        </Field>
    );
};
