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
| `Excel公式经验库.md` | Markdown 版本经验库 |
| `.nojekyll` | 让 GitHub Pages 原样托管静态文件 |

## 更新方式

后续新增公式场景时，主要更新：

```text
web/index.html
Excel公式经验库.md
```

