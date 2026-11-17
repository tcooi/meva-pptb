import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import type { Plugin } from 'vite';

/**
 * Plugin to fix HTML for PPTB compatibility
 * - Removes type="module" and crossorigin attributes since we're using IIFE format
 * - Moves script tags from head to end of body so DOM is ready when IIFE executes
 * - In builds, loads the bundle only after the window `load` event (see below)
 */
function fixHtmlForPPTB(): Plugin {
    return {
        name: 'fix-html-for-pptb',
        enforce: 'post',
        transformIndexHtml(html, ctx) {
            // Remove type="module" and crossorigin from script tags
            // IIFE format doesn't need module type, and file:// URLs don't need crossorigin
            html = html.replace(/\s*type="module"/g, '');
            html = html.replace(/\s*crossorigin/g, '');
            // Clean up extra spaces around attributes
            html = html.replace(/\s+>/g, '>');

            // Move script tags from head to end of body
            // IIFE executes immediately, so DOM must be ready
            const scriptRegex = /(<script[^>]*src="[^"]*"[^>]*><\/script>)/g;
            const scripts: string[] = [];

            // Extract all script tags
            html = html.replace(scriptRegex, (match) => {
                scripts.push(match);
                return ''; // Remove from current position
            });

            // PPTB keeps a tool hidden until its page fires `load`, which normally waits for the
            // whole bundle to download and run. Injecting the bundle after `load` lets PPTB show
            // the static loading screen in index.html straight away. Dev server keeps normal tags.
            if (!ctx.server && scripts.length > 0) {
                const sources = scripts.map((tag) => tag.match(/src="([^"]*)"/)![1]);
                const loader = `<script>
    window.addEventListener('load', function () {
      ${JSON.stringify(sources)}.forEach(function (src) {
        var script = document.createElement('script');
        script.src = src;
        script.async = false;
        document.body.appendChild(script);
      });
    });
  </script>`;
                scripts.length = 0;
                scripts.push(loader);
            }

            // Insert scripts before closing body tag
            if (scripts.length > 0) {
                const scriptsHtml = '\n  ' + scripts.join('\n  ');
                html = html.replace('</body>', scriptsHtml + '\n</body>');
            }

            return html;
        },
    };
}

// https://vitejs.dev/config/
export default defineConfig((configEnv) => {
    return {
        plugins: [react(), fixHtmlForPPTB()],
        base: './',
        build: {
            outDir: 'dist',
            assetsDir: 'assets',
            sourcemap: configEnv.mode === 'development',
            rollupOptions: {
                output: {
                    // Use IIFE format for compatibility with iframe srcdoc loading
                    // ES modules can have issues when loaded via file:// URLs in iframes
                    format: 'iife',
                    // Bundle everything into a single file to avoid module loading issues
                    inlineDynamicImports: true,
                    // Disable chunking since we're bundling everything
                    manualChunks: undefined,
                },
            },
        },
    };
});
