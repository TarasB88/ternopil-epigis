import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: "dist",
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (!id.includes("node_modules")) return;
          if (id.includes("react") || id.includes("scheduler")) return "vendor-react";
          if (id.includes("leaflet")) return "vendor-map";
          if (id.includes("@turf")) return "vendor-geo";
          if (id.includes("xlsx")) return "vendor-spreadsheet";
          if (id.includes("jspdf") || id.includes("html2canvas") || id.includes("canvg") || id.includes("dompurify")) return "vendor-export";
          if (id.includes("lucide-react")) return "vendor-icons";
          return "vendor";
        }
      }
    }
  }
});
