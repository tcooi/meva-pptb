import { useCallback, useEffect, useRef, useState } from 'react';

const SETTINGS_KEY = 'ui.columnWidths';
export const MIN_COLUMN_WIDTH = 60;

/**
 * Column widths keyed by column id, persisted in the tool's PPTB settings so they survive restarts.
 * Ids without a stored width fall back to the caller's default.
 */
export function useColumnWidths() {
    const [widths, setWidths] = useState<Record<string, number>>({});
    const saveTimer = useRef<ReturnType<typeof setTimeout>>();

    useEffect(() => {
        (async () => {
            try {
                const stored = await window.toolboxAPI.settings.get(SETTINGS_KEY);
                if (stored && typeof stored === 'object') {
                    // Keep anything the user already resized before the stored widths arrived
                    setWidths((current) => ({ ...(stored as Record<string, number>), ...current }));
                }
            } catch (error) {
                console.error('Error loading column widths:', error);
            }
        })();
    }, []);

    const persist = useCallback((next: Record<string, number>) => {
        // Debounced: dragging fires many updates, only the final width needs saving
        clearTimeout(saveTimer.current);
        saveTimer.current = setTimeout(async () => {
            try {
                await window.toolboxAPI.settings.set(SETTINGS_KEY, next);
            } catch (error) {
                console.error('Error saving column widths:', error);
            }
        }, 400);
    }, []);

    const setWidth = useCallback(
        (id: string, width: number | undefined) => {
            setWidths((current) => {
                const next = { ...current };
                if (width === undefined) {
                    delete next[id];
                } else {
                    next[id] = Math.max(MIN_COLUMN_WIDTH, Math.round(width));
                }
                persist(next);
                return next;
            });
        },
        [persist]
    );

    return { widths, setWidth };
}
