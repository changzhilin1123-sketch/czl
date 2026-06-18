# Nginx-RTMP · 无人机局域网中继

> 武隆直播：无人机 RTMP → nginx → VLC → 抖音直播伴侣窗口采集。

## 便携包（复制到任意工作电脑）

**请使用 [`操作手册-便携版.md`](./操作手册-便携版.md)**，复制整个 `nginx-rtmp` 文件夹即可。

| 步骤 | 操作 |
|------|------|
| 首次 | 双击 `1-安装.bat` |
| 启动 | `2-启动.bat` |
| 验证 | `3-验证.bat`（须 PASS=5） |
| 停止 | `4-停止.bat` |

默认安装到 `nginx-rtmp\runtime\`；本机已有安装可复制 `scripts\paths.local.bat.example` → `paths.local.bat` 指定目录。

## 项目内文档

- 业务上下文：`../设备租赁上下文.md` §5.6
- 配置：`conf/nginx.conf`
- 脚本：`scripts/`（由根目录 `1-4.bat` 调用）

## 下载源

https://github.com/iliweii/nginx-rtmp-win64/releases
