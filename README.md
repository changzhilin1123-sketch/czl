# 武隆景区

本工作区用于 **武隆景区** 相关事项（直播、运营、内容等）。

- **总入口**：[`项目上下文.md`](./项目上下文.md)
- **直播设备**：[`直播设备/设备租赁上下文.md`](./直播设备/设备租赁上下文.md)

---

## 跨电脑同步说明

**GitHub 仓库**：<https://github.com/changzhilin1123-sketch/czl.git>  
**本地路径**（建议两台电脑保持一致）：`d:\cursor\武隆景区`

### 另一台电脑首次使用

1. 安装 [Git for Windows](https://git-scm.com/download/win)
2. 克隆仓库：

```powershell
git clone https://github.com/changzhilin1123-sketch/czl.git "d:\cursor\武隆景区"
```

3. Cursor → **File → Open Folder** → 选择 `d:\cursor\武隆景区`

### 日常同步

| 场景 | 操作 |
|------|------|
| 在本机改完文件 | `git add .` → `git commit -m "简要说明"` → `git push` |
| 换电脑开工前 | 先 `git pull`，再打开 Cursor 编辑 |
| 两边都改过 | 先 `git pull`，有冲突按提示合并后再 `git push` |

所有命令在 `d:\cursor\武隆景区` 目录下执行。

### 会同步 / 不会同步

| 会随 Git 同步 | 不会自动同步 |
|---------------|--------------|
| `项目上下文.md`、子目录文档、Excel、脚本 | Cursor **聊天记录** |
| `.cursor/rules/`（AI 工作规则） | 各电脑单独的 Cursor 全局设置 |

重要结论请写入 `项目上下文.md` 变更日志，不要依赖聊天历史。

### 常见问题

- **终端找不到 `git` 命令**：新开终端，或使用完整路径 `D:\Git\Git\cmd\git.exe`（以本机 Git 安装位置为准）。
- **推送时要登录 GitHub**：按提示在浏览器完成 GitHub 账号授权即可。
