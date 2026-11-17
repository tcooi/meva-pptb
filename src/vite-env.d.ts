/// <reference types="@pptb/types" />

declare global {
    interface Window {
        toolboxAPI: typeof import('@pptb/types').toolboxAPI;
        dataverseAPI: typeof import('@pptb/types').dataverseAPI;
        /** Theme request started by the inline script in index.html. */
        __pptbTheme?: Promise<'light' | 'dark'>;
    }
}

export {};
