import { useEffect, useState } from 'react';
import { describeError, loadUnmanagedSolutions, type UnmanagedSolution } from '../services/environmentVariables';

/** "No solution" option in the solution pickers. */
export const NO_SOLUTION = '';

/**
 * Loads the unmanaged solutions (and their publisher prefixes) of one environment while a dialog is open,
 * and preselects one: the first of `preferred` that exists there, else the environment's default solution.
 */
export function useUnmanagedSolutions(open: boolean, slot: number | undefined, preferred: string[] = []) {
    const [solutions, setSolutions] = useState<UnmanagedSolution[] | null>(null);
    const [error, setError] = useState<string>();
    // Kept separate from `solutions` so the previous list (and prefix) stays put while another environment loads
    const [loading, setLoading] = useState(true);
    const [selected, setSelected] = useState(NO_SOLUTION);
    const preferredKey = preferred.join('\n');

    useEffect(() => {
        if (!open || slot === undefined) return;
        let cancelled = false;
        setLoading(true);
        setError(undefined);
        loadUnmanagedSolutions(slot)
            .then((list) => {
                if (cancelled) return;
                setSolutions(list);
                setLoading(false);
                // Otherwise the environment's default unmanaged solution when there is one, like the maker portal
                const match =
                    preferred.map((name) => list.find((s) => s.uniqueName === name)).find(Boolean) ??
                    list.find((s) => s.friendlyName === 'Common Data Services Default Solution');
                setSelected(match?.uniqueName ?? list[0]?.uniqueName ?? NO_SOLUTION);
            })
            .catch((err) => {
                if (cancelled) return;
                console.error('Error loading solutions:', err);
                setSolutions([]);
                setSelected(NO_SOLUTION);
                setError(describeError(err));
                setLoading(false);
            });
        return () => {
            cancelled = true;
        };
    }, [open, slot, preferredKey]);

    const solution = solutions?.find((s) => s.uniqueName === selected) ?? null;
    return { solutions, loading, error, selected, setSelected, solution };
}
