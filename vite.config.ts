import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Fully client-side app. Base is relative so the built /dist works from any
// static host or opened via a sub-path.
export default defineConfig({
  base: "./",
  plugins: [react()],
  build: {
    target: "es2022",
    chunkSizeWarningLimit: 2000,
  },
  worker: {
    format: "es",
  },
});
