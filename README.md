# CheckSentry

CheckSentry 是一个纯本地运行的 Windows IT 合规核对工具。它从 Windows 卸载注册表读取已安装软件，扫描主流 Chromium 浏览器及 Firefox 插件，并与 `list.xlsx` 中的三类规则清单核对。

工具只负责发现、展示和维护清单，不会自动卸载软件。卸载仍由用户在 Revo Uninstaller 等工具中人工完成。

## 主要功能

- 显式读取 Windows 32 位和 64 位注册表视图，不调用 `Win32_Product`。
- 扫描 Chrome、Edge、Chromium、Brave、Vivaldi、Naver Whale 和 Firefox。
- 使用 `list.xlsx` 作为唯一读写数据源，不使用数据库或 Office COM。
- 固定匹配优先级：**黑名单 → 白名单 → 待定**。
- 未命中任何规则的软件或用户插件会自动加入“待定”。
- 本地报告页支持区域全选、批量分类、TSV 复制、图标、安装位置/个人资料和版本变化提示。
- `/manage` 页面支持三类清单切换、视图排序、直接编辑、新增、跨表移动、删除和单步撤销。
- Excel 写入采用临时副本校验后提交；撤销前会检测外部修改，避免覆盖同时发生的 Excel 编辑。
- 本地写接口使用启动期随机安全令牌、同源检查、请求大小限制和服务端字段验证。

## 文件说明

| 文件 | 用途 |
|---|---|
| `运行核对工具.bat` | Windows PowerShell 日常启动入口 |
| `Start-ComplianceCheck.ps1` | 核心核对、Excel 事务写入和本地 HTTP 服务 |
| `Get-InstalledSoftware.ps1` | 软件注册表扫描 |
| `Get-InstalledExtensions.ps1` | 七类浏览器插件扫描 |
| `report_template.html` | 报告页面模板 |
| `management_template.html` | 清单维护页面模板 |
| `list_template.xlsx` | 四张工作表的空白模板 |
| `TESTING.md` | Windows 虚拟机发布前验收清单 |
| `list.xlsx` | 运行后生成的个人清单；已被 `.gitignore` 排除，禁止上传 GitHub |

## 系统要求

- Windows 10/11 或相应 Windows Server。
- Windows PowerShell 5.1，或 PowerShell 7。
- 首次运行时访问 PowerShell Gallery，以安装经过验证的 `ImportExcel 7.8.10`。
- 不需要安装 Microsoft Office。

如果目标环境不能联网，请在可联网环境下载并按企业软件分发流程安装 `ImportExcel 7.8.10`，然后再运行 CheckSentry。

## 首次运行

1. 解压完整发布包到本地普通目录。
2. 双击 `运行核对工具.bat`。
3. 工具在 `list.xlsx` 不存在时，从空白 `list_template.xlsx` 复制创建。
4. 默认尝试监听 `http://localhost:8787/`。如果 8787 被占用，会尝试后续端口，并在命令窗口显示实际地址。
5. 首次扫描到的项目会自动写入“待定”，然后显示在报告中。

使用 PowerShell 7 时可以运行：

```powershell
pwsh -NoProfile -File .\Start-ComplianceCheck.ps1
```

指定端口或清单路径：

```powershell
powershell.exe -NoProfile -File .\Start-ComplianceCheck.ps1 -Port 8790 -ListPath "D:\CheckSentryData\list.xlsx"
```

清单路径必须是现有目录中的 `.xlsx` 文件，且不能直接指向 `list_template.xlsx`。

## 固定匹配优先级

1. **黑名单最高**：一旦命中，立即标红，不再检查白名单或待定。
2. **白名单次之**：未命中黑名单时，命中白名单则显示“已匹配”或“版本变化”。
3. **待定最低**：仅在前两类都未命中时使用。

同一对象即使同时存在于多张工作表，程序也不会在启动时擅自删除规则，并始终按以上顺序判断。

## 匹配方式

- `插件ID精确`：插件唯一 ID 完全相同。
- `精确`：名称完全相同，不区分大小写。
- `包含`：按普通文本包含关系匹配；`*`、`?`、`[` 不会被当作通配符。
- `通配符`：显式支持 PowerShell 通配符语法。
- `正则`：使用带执行超时的正则表达式；无效或高耗时表达式会被拒绝。

版本号留空表示不限制；白名单版本可使用 `128.*` 等通配符。

## 网页维护和撤销

- 网页排序只改变当前显示顺序，不会改变 Excel 行顺序或逐行匹配顺序。
- 每次新增、编辑、删除或批量分类会保存一个完整工作簿快照。
- 只有最近一次网页维护操作可以撤销。
- 如果上一步操作后 `list.xlsx` 被 Excel、同步软件或其他进程修改，撤销会被拒绝，避免覆盖外部修改。
- 扫描时自动加入待定不提供撤销，并会使旧撤销快照失效。

## Excel 文件占用和异常

- 工具运行时不要同时用 Excel 编辑 `list.xlsx`。
- 文件被占用、缺少工作表、列名被修改或规则单元格包含公式时，工具会停止写入并显示错误。
- 写入在临时副本中完成，结构验证通过且原文件未被外部修改后才提交。

## 安全说明

- HTTP 服务只绑定 `localhost`，不会监听局域网地址。
- 所有写接口要求页面中的随机安全令牌，并检查同源、JSON Content-Type 和请求大小。
- 用户字段按纯文本写入 Excel；不执行清单中的 PowerShell、HTML、JavaScript 或 Excel 公式。
- 图标只读取已扫描到的本地文件，限制大小并验证常见位图文件头；不加载 SVG 或远程 URL。
- `list.xlsx` 包含软件、插件、用户名和审批备注等信息，禁止上传公开仓库。

## GitHub 发布清单

正式发布包应只包含：

- `README.md`
- `TESTING.md`
- `运行核对工具.bat`
- `Start-ComplianceCheck.ps1`
- `Get-InstalledSoftware.ps1`
- `Get-InstalledExtensions.ps1`
- `report_template.html`
- `management_template.html`
- `list_template.xlsx`

发布前确认包内没有 `list.xlsx`、`.DS_Store`、`~$*.xlsx`、备份、日志、测试结果或临时文件。

## 如何发布新版本

本项目只允许通过 GitHub Actions 自动构建、生成 Artifact Attestation 并上传 Release。不要在 GitHub 网页中手工创建 Release、上传或替换文件，否则构建证明会失效。

### 1. 确保代码已经提交并推送

```bash
git status
git add .
git commit -m "你的改动说明"
git push origin main
```

### 2. 创建版本 Tag

Tag 必须使用 `v主版本.次版本.修订版本` 格式，例如：

```bash
git tag -a v1.0.1 -m "Release version 1.0.1"
```

### 3. 推送 Tag，触发自动发布

```bash
git push origin v1.0.1
```

GitHub Actions 会在同一个任务中完成源文件校验、PowerShell 5.1/7 语法检查、Excel 模板检查、最终 ZIP 构建、SHA-256 生成、构建证明签名和 Release 上传。

构建证明的 `subject-path` 与 Release 上传路径引用同一份最终文件。签名前后还会再次核对 SHA-256，避免对中间文件签名或签名后重新压缩导致哈希漂移。

### 4. 验证构建结果

- 在仓库的 **Actions** 页面确认 `Build, Attest and Release` 成功。
- 在 **Releases** 页面确认 ZIP 和 `.sha256` 均由 `github-actions[bot]` 上传。
- 下载 ZIP 后执行：

```bash
gh attestation verify CheckSentry-v1.0.1.zip --repo secure-artifacts/CheckSentry
```

### 版本号说明

| 格式 | 用途 | 示例 |
|---|---|---|
| `vX.0.0` | 重大或不兼容更新 | `v2.0.0` |
| `vX.Y.0` | 新功能 | `v1.1.0` |
| `vX.Y.Z` | 错误或安全修复 | `v1.0.1` |

### 构建失败时

先查看 Actions 日志并修复代码或 Workflow，然后删除失败的 Tag 并重新创建。不要人工补传 Release 文件。

```bash
git tag -d v1.0.1
git push origin :refs/tags/v1.0.1
git tag -a v1.0.1 -m "Release version 1.0.1"
git push origin v1.0.1
```
