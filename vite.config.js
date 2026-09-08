import { defineConfig } from "vite";

// GitHub Pages serves project sites from /<repo-name>/, so asset URLs
// need that prefix baked in at build time.
export default defineConfig({
  base: "/referral-form/",
});
