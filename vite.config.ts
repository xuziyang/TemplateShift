import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import electron from 'vite-plugin-electron'
import { resolve } from 'path'
import { rmSync } from 'fs'

rmSync('dist-electron', { recursive: true, force: true })

export default defineConfig({
  plugins: [
    react(),
    electron([
      {
        entry: 'src/main/main.ts',
        vite: {
          build: {
            outDir: 'dist-electron',
            rollupOptions: {
              external: ['electron']
            }
          }
        }
      }
    ]),
  ],
  resolve: {
    alias: {
      '@': resolve(__dirname, './src'),
    },
  },
  base: process.env.ELECTRON_RENDERER_URL ? '/' : './',
  build: {
    outDir: 'dist/renderer',
    emptyOutDir: true,
    assetsDir: 'assets',
  },
})