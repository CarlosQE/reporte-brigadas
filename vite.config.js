import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Cambia 'reporte-brigadas' por el nombre exacto de tu repositorio en GitHub
export default defineConfig({
  plugins: [react()],
  base: '/reporte-brigadas/',
})
