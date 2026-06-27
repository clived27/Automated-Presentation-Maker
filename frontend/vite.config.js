import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      // Forward /api/* to the local Python backend (http.server on port 8000)
      '/api': {
        target: 'http://localhost:8000',
        changeOrigin: true,
        // Strip the /api prefix — Python handler listens on /generate-ppt
        rewrite: (path) => path.replace(/^\/api/, ''),
      },
    },
  },
})
