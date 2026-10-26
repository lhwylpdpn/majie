# 京购商城 Demo

对标京东首页的极简纯静态电商演示页，用于思路验证与原型展示。

## 目录结构

```
demo/
  index.html      # 首页
  cart.html       # 购物车
  css/style.css   # 样式
  js/
    config.js     # 站点与商品配置
    main.js       # 首页逻辑
    cart.js       # 购物车逻辑
```

## 功能概览

- 顶栏、Logo、搜索框、购物车入口
- 左侧分类菜单 + 轮播 Banner + 右侧快捷入口
- 秒杀区（倒计时）、推荐商品、热销榜单
- 加入购物车（localStorage 持久化）
- 购物车页：数量调整、全选、合计、结算提示

## 本地预览

```bash
cd demo
python3 -m http.server 8080
```

浏览器访问 http://localhost:8080

也可直接用浏览器打开 `index.html`（部分浏览器对 localStorage 跨文件访问有限制，推荐用本地服务器）。

## 配置修改

编辑 `js/config.js` 中的 `SITE_CONFIG` 与 `PRODUCTS` 数组即可调整站点名称、主题色和商品数据。
