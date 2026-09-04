地基雷达监测报告生成器 Mobile V4
================================

V4 依据最新桌面版 wjz shibao_zhong.py 审查并修正后制作。

核心功能：
1. 短时面 + 短时点：生成 6 小时报告。
2. 长时面 + 长时点：生成 24 小时报告和周报。
3. 24 小时日报正文只显示当前 24 小时监测结果，不再输出“与上周相比”。
4. 24 小时页面仍保留“Excel 对比周期”，仅用于导出的 24h 对比表。
5. 周报继续输出“与上周相比”的位移变化。
6. 修正 Excel“总累积位移量”：直接读取期末累计值，而不是错误地用“期末-首期”。
7. 修正无对比周期时 Excel 表头/数据列数不一致的问题。
8. 改进 6h 自动推荐，按真实 6 小时时间桶选择，避免跨日期误选。
9. 增加 M1-M14 / S1-S14 必要字段检查，缺列时直接提示。
10. 保留 nodim:true，兼容工作表 !ref 范围错误的雷达 Excel。

文件：
- index.html              手机/PWA入口
- app.js                  V4核心逻辑
- styles.css              手机样式
- manifest.webmanifest    Web App配置
- sw.js                   V4离线缓存/旧缓存清理
- .nojekyll               GitHub Pages静态发布
- icon-192.png / icon-512.png
- wjz_shibao_zhong_V4_fixed.py  同步修正后的桌面Python版
- V4修正说明.txt
- 完整设置教程.md

使用方法与 V3 一致：上传到原 GitHub Pages 仓库根目录覆盖 V3，等待 Pages 部署完成，在 Safari 强制刷新；如主屏幕仍显示 V3，删除旧 Web App 后重新“添加到主屏幕”。
