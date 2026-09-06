地基雷达监测报告生成器 Mobile V4.2 CLEAN

本包用于修复“上传 V4.1 后 GitHub Pages 反而显示 V2”的部署混淆问题。
核心策略：index.html + 全新的版本化资源文件名，避免旧 app.js/styles.css/sw.js 与浏览器缓存混用。

上传到 GitHub 仓库根目录的文件：
- index.html（必须覆盖旧 index.html）
- app-v4.2.js
- styles-v4.2.css
- sw-v4.2.js
- manifest-v4.2.webmanifest
- icon-192.png
- icon-512.png
- .nojekyll

旧的 app.js / styles.css / sw.js / manifest.webmanifest 可以暂时保留；V4.2 不再引用它们。
部署后页面顶部必须显示：V4.2 · 双数据源 · 6小时 / 24小时 / 周报
并应看到 4 个文件入口：短时面、短时点、长时面、长时点。
