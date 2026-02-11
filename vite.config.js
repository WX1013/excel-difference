import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { cloudflare } from '@cloudflare/vite-plugin'

export default defineConfig({
  plugins: [
    vue(),           // Vue 插件
    cloudflare()     // Cloudflare Vite 插件
  ],
  build: {
    outDir: 'dist'   // 与 wrangler.toml 中的 dist_dir 一致
  }
})
