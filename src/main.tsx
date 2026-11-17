import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import './index.css';

/**
 * Resolves PPTB's theme before the first render so React never paints a light frame in dark mode.
 * Reuses the request index.html already started; falls back to the system setting if PPTB is slow.
 */
async function resolveInitialTheme(): Promise<'light' | 'dark'> {
    const systemTheme = window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    try {
        const request = window.__pptbTheme ?? window.toolboxAPI?.utils.getCurrentTheme();
        if (!request) return systemTheme;
        const timeout = new Promise<'light' | 'dark'>((resolve) => setTimeout(() => resolve(systemTheme), 500));
        return (await Promise.race([request, timeout])) === 'dark' ? 'dark' : 'light';
    } catch (error) {
        console.error('Error getting theme:', error);
        return systemTheme;
    }
}

// Ensure DOM is ready and root element exists
const rootElement = document.getElementById('root');
if (rootElement && !rootElement.hasAttribute('data-reactroot-initialized')) {
    // Mark as initialized to prevent double rendering
    rootElement.setAttribute('data-reactroot-initialized', 'true');

    resolveInitialTheme().then((initialTheme) => {
        createRoot(rootElement).render(
            <StrictMode>
                <App initialTheme={initialTheme} />
            </StrictMode>
        );
    });
} else if (!rootElement) {
    console.error('Root element not found. Make sure the HTML contains <div id="root"></div>');
}
