import { defineConfig } from 'vite';
import basicSsl from '@vitejs/plugin-basic-ssl';

// HTTPS (certificado autofirmado) es obligatorio para que el celular
// permita usar la cámara cuando se accede por la red local (192.168.x.x).
export default defineConfig({
  base: './',
  plugins: [basicSsl()],
  server: { host: true, port: 5173 },
});
