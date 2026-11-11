import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    // `npm run dev` proxies API calls to a running `wrangler dev` (port 8787)
    proxy: { "/api": "http://127.0.0.1:8787" },
  },
  build: { chunkSizeWarningLimit: 2000 },
});
