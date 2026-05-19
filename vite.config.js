import { defineConfig } from 'vite';
import { resolve } from 'path';

export default defineConfig({
  root: 'src',
  build: {
    outDir: '../dist',
    emptyOutDir: true,
    rollupOptions: {
      input: {
        main: resolve(__dirname, 'src/index.html'),
        punto: resolve(__dirname, 'src/punto-de-equilibrio/index.html'),
        servicios: resolve(__dirname, 'src/servicios-contables.html'),
        legal: resolve(__dirname, 'src/asesoria-legal.html'),
        gestion: resolve(__dirname, 'src/gestion-financiera.html'),
      }
    }
  }
});
