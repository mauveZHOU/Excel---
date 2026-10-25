# Excel 公式经验库

这是一个用于沉淀日常数据分析公式经验的静态网页知识库。

## 直接打开

下载仓库后，双击打开：

```text
index.html
```

或者打开：

```text
web/index.html
```

## GitHub Pages 发布

1. 把本文件夹上传到一个 GitHub 仓库。
2. 进入 GitHub 仓库的 `Settings`。
3. 打开 `Pages`。
4. Source 选择 `Deploy from a branch`。
5. Branch 选择 `main`，目录选择 `/root`。
6. 保存后等待 GitHub 生成访问链接。

发布后，其他电脑可以直接通过 GitHub Pages 链接打开经验库。

## 文件说明

| 文件/目录 | 作用 |
|---|---|
| `index.html` | GitHub Pages 首页入口，会跳转到主网页 |
| `web/index.html` | 主要网页版经验库 |
| `web/practice.html` | 练习区：工作日每天自动出一组题，在网页里写公式、自动判对错 |
| `web/practice/` | 练习区的公式引擎（engine.js）、题库（bank.js）、判题（core.js）和页面逻辑（app.js） |
| `Excel公式经验库.md` | Markdown 版本经验库 |
| `.nojekyll` | 让 GitHub Pages 原样托管静态文件 |

## 更新方式

后续新增公式场景时，主要更新：

```text
web/index.html
Excel公式经验库.md
```


## 练习区

打开 `web/practice.html`（经验库左侧目录里的「练习区」）。

- **每日推送**：周一到周五，每天第一次打开时自动准备一组题（3 或 5 道）。做错、看过答案的题第二天再出现；做对的题按 1、2、4、7、15、30 天的间隔回来复习。
- **三种题**：表格题（在网页模拟的表格里写公式，往下/往右自动填充、逐格判对错）、回忆题、填空题。
- **自动跟着经验库更新**：练习区会读取 `web/index.html` 里的「原公式」和「主动遮挡自测」，自动生成填空题和回忆题。新增场景后不用改练习区，新公式自动进入题库。
- **表格题**写在 `web/practice/bank.js`，新增场景时可以按同样的格式补充。
- **进度保存在浏览器里**：换电脑或清理浏览器前，在练习区「备份进度」里导出 .json 文件。
- 直接双击打开本地文件时读不到经验库页面，只有表格题；用 GitHub Pages 网址打开时三种题都有。
