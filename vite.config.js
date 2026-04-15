import { defineConfig } from 'vite';
import tailwindcss from '@tailwindcss/vite';

export default defineConfig({
  base: '/slide-color-changer/',
  plugins: [tailwindcss()],
});
