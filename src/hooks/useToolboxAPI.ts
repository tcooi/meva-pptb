import { useCallback, useEffect, useRef, useState } from 'react';

/**
 * Returns the tool's connection slots. Null gaps are preserved so the array index
 * is always the connection target to pass to dataverseAPI.
 */
export function useConnections() {
    const [connections, setConnections] = useState<(ToolBoxAPI.Connection | null)[]>([]);
    const [isLoading, setIsLoading] = useState(true);

    const refreshConnections = useCallback(async () => {
        try {
            setConnections(await window.toolboxAPI.connections.getConnections());
        } catch (error) {
            console.error('Error loading connections:', error);
        } finally {
            setIsLoading(false);
        }
    }, []);

    useEffect(() => {
        refreshConnections();
    }, [refreshConnections]);

    return { connections, isLoading, refreshConnections };
}

export function useTheme(initialTheme: 'light' | 'dark') {
    const [theme, setTheme] = useState<'light' | 'dark'>(initialTheme);

    const refreshTheme = useCallback(async () => {
        try {
            setTheme((await window.toolboxAPI.utils.getCurrentTheme()) === 'dark' ? 'dark' : 'light');
        } catch (error) {
            console.error('Error getting theme:', error);
        }
    }, []);

    useEffect(() => {
        refreshTheme();
    }, [refreshTheme]);

    // Keeps index.html's color-scheme rules (scrollbars, page canvas) in step with PPTB's theme
    useEffect(() => {
        document.documentElement.setAttribute('data-theme', theme);
    }, [theme]);

    return { theme, refreshTheme };
}

/**
 * Subscribes a single handler to ToolBox events for the lifetime of the component.
 * The latest callback is kept in a ref so re-renders don't re-register the handler.
 */
export function useToolboxEvents(onEvent: (event: ToolBoxAPI.ToolBoxEvent, data: unknown) => void) {
    const callbackRef = useRef(onEvent);
    callbackRef.current = onEvent;

    useEffect(() => {
        const handler = (_event: unknown, payload: ToolBoxAPI.ToolBoxEventPayload) => {
            try {
                callbackRef.current(payload.event, payload.data);
            } catch (error) {
                console.error('Error handling ToolBox event:', error);
            }
        };

        window.toolboxAPI.events.on(handler);
        return () => window.toolboxAPI.events.off(handler);
    }, []);
}

export async function notify(title: string, body: string, type: ToolBoxAPI.NotificationOptions['type']) {
    try {
        await window.toolboxAPI.utils.showNotification({ title, body, type, duration: type === 'error' ? 0 : 3000 });
    } catch (error) {
        console.error('Error showing notification:', error);
    }
}
