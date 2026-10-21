import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import { fileURLToPath } from 'node:url';

export default defineConfig({
  plugins: [react(), tailwindcss()],
  build: {
    outDir: process.env.VERCEL ? 'dist' : fileURLToPath(new URL('./backend/static', import.meta.url)),
    emptyOutDir: true,
  },
});
