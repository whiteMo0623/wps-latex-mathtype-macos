# WPS LaTeX ↔ MathType for macOS

在 macOS 版 WPS 文字里，批量将选区中的 LaTeX 转为 MathType 公式对象，也能把选中的 MathType 对象转回 LaTeX。转换直接修改打开的 `.docx`。应用还支持逐文件将 LaTeX 转为 MathType，结果保存为副本。

## 安装

1. 到 [Releases](../../releases) 下载对应芯片的 ZIP。Apple 芯片选 `arm64`，Intel 芯片选 `x86_64`。应用的目标系统是 macOS 13 或更新版本。
2. 解压后将 **WPS LaTeX MathType.app** 拖入「应用程序」。首次使用前先打开一次 WPS Office。
3. 打开应用，点 **安装或更新 WPS 加载项**。如果 macOS 弹出文件夹选择窗口，在已定位的目录点 **选择此文件夹**。
4. 保存工作并彻底退出 WPS，再打开一份 `.docx`。按提示信任名称为 `LaTeXToMathTypeSelection`、地址为 `http://127.0.0.1:49265/` 的加载项。

ZIP 内含 Python、Node.js 及转换依赖。安装和转换均可离线进行。WPS Office 由用户自行安装。

当前发行包使用 ad hoc 签名，尚未经过 Apple 公证。若 macOS 拦截首次运行，确认下载来源和校验值后，可在「系统设置 → 隐私与安全性」为该应用选择「仍要打开」。[Apple 的说明](https://support.apple.com/en-us/102445)

## 在 WPS 中使用

- 选中一段文字，或按 **Command+A** 扩大到整篇正文。打开 **LaTeX / MathType** 功能区，点 **LaTeX → MathType**。
- 单击一个 MathType 公式对象，或选中包含多个对象的范围，再点 **MathType → LaTeX**。
- 逐文件转换可在应用中点 **从文件转换 LaTeX → MathType**。选择一份或多份 `.docx`，转换后的文件以 `_MathType7.docx` 为后缀，保存在源文件旁。

可以在正文或表格内测试下面两种公式。也支持 `\(…\)` 和 `$$…$$`。

```text
行内公式 $E=mc^2$ 位于段落中。
\[\frac{a+b}{c}\]
```

### 兼容性

- WPS 编辑器内的双向转换已在 Apple 芯片 Mac 的 **WPS Office 12.1.26026** 验证。加载项依赖 WPS 的 JS 加载项支持以及 `Range.InsertXML` 和 `Range.WordOpenXML`。Intel 发行包的转换引擎已通过 Rosetta 测试，Intel Mac 上的 WPS 编辑器仍待实机验证。
- 输出对象为 `Equation.DSMT4` OLE，内含 **Equation Native** 数据。已验证的 Mac WPS 版本无法双击启动 MathType 7 桌面编辑器。在 Mac 上可将对象转回 LaTeX 再修改。在安装了 MathType 7 的 Windows Word 中可使用桌面编辑器处理兼容对象。[WIRIS 的 MathType 7 系统要求](https://docs.wiris.com/installing-mathtype-7-for-windows-and-mac)
- 反向转换读取 MTEF v5，生成的 LaTeX 可能采用等价写法。选区内跨段的公式、旧版 MTEF、修订或域等无法安全替换的内容会保留。失败项和原因由加载项报告。

### 更新与卸载

更新时在原位置替换应用，再点 **安装或更新 WPS 加载项**。移动应用后也需重新安装加载项，因为后台服务记录应用的新路径。卸载时先点 **卸载 WPS 加载项**，然后重启 WPS，再删除应用。

本地转换服务由当前用户的 LaunchAgent 启动，仅监听 `127.0.0.1:49265`。启动日志位于 `~/Library/Logs/LaTeXToMathType-WPS.log`。若 WPS 的组织策略禁止 JS 加载项，需在允许加载项的 WPS 环境中使用编辑器集成。

## 从源码开发

需要 macOS、Python 3.12 和 Node.js 22。文件对象生成、选区转换、本地服务和 WPS 功能区位于 `src/`。应用界面位于 `macos/Main.swift`。

```sh
python3.12 -m venv .venv
.venv/bin/python -m pip install --only-binary=:all: --require-hashes -r requirements.lock
npm ci --ignore-scripts
.venv/bin/python -B -m unittest discover -s tests -v
```

从源码运行文件转换时，可传入一个或多个 `.docx` 文件。

```sh
.venv/bin/python -B src/convert.py "/path/to/document.docx"
```

## 构建发行包

安装 Xcode Command Line Tools 后，在 macOS 上运行下面的命令。构建时会下载上游运行时和依赖并校验哈希。每个 ZIP 都会在新路径下解压，并在空白用户目录中运行选区转换和文件转换测试。构建另一架构时，本机需能执行该架构的二进制。

```sh
python3 scripts/build-release.py --arch arm64
python3 scripts/build-release.py --arch x86_64
```

`dist/` 含两个 ZIP 和对应的 `.sha256` 文件。将这些文件上传为 GitHub Release 附件。仓库中的 Actions 工作流也可以分别构建两个架构并生成产物。版本号维护在 `package.json`，发布时应同步更新 `package-lock.json` 并使用同一版本的 Git 标签。

源码采用 [MIT 许可证](LICENSE)。随发行包分发的第三方组件保留各自的许可和署名。[第三方组件说明](THIRD_PARTY_NOTICES.md)
