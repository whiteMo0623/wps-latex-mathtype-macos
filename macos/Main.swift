import AppKit
import UniformTypeIdentifiers

final class Controller: NSObject, NSApplicationDelegate, NSWindowDelegate {
    private var window: NSWindow!
    private var buttons: [NSButton] = []
    private var busy = false

    func applicationDidFinishLaunching(_ notification: Notification) {
        let quit = NSMenuItem(
            title: "退出 WPS LaTeX ↔ MathType",
            action: #selector(NSApplication.terminate(_:)),
            keyEquivalent: "q"
        )
        let menu = NSMenu()
        let appMenu = NSMenuItem()
        appMenu.submenu = NSMenu()
        appMenu.submenu?.addItem(quit)
        menu.addItem(appMenu)
        let editMenu = NSMenuItem()
        editMenu.submenu = NSMenu(title: "编辑")
        let edits = [
            ("剪切", "cut:", "x"),
            ("复制", "copy:", "c"),
            ("粘贴", "paste:", "v"),
            ("全选", "selectAll:", "a"),
        ]
        for (title, action, key) in edits {
            editMenu.submenu?.addItem(withTitle: title, action: Selector(action), keyEquivalent: key)
        }
        menu.addItem(editMenu)
        NSApp.mainMenu = menu

        window = NSWindow(
            contentRect: NSRect(x: 0, y: 0, width: 400, height: 210),
            styleMask: [.titled, .closable, .miniaturizable],
            backing: .buffered,
            defer: false
        )
        window.title = "WPS LaTeX ↔ MathType"
        window.delegate = self
        window.isReleasedWhenClosed = false

        let actions: [(String, Selector)] = [
            ("安装或更新 WPS 加载项", #selector(install)),
            ("从文件转换 LaTeX → MathType", #selector(convert)),
            ("卸载 WPS 加载项", #selector(uninstall)),
        ]
        buttons = actions.map { title, action in
            NSButton(title: title, target: self, action: action)
        }
        let note = NSTextField(wrappingLabelWithString: "在 WPS 中打开 .docx，可对选区进行双向转换。")
        note.alignment = .center
        let layout = NSStackView(views: [note] + buttons)
        layout.orientation = .vertical
        layout.alignment = .centerX
        layout.spacing = 12
        layout.translatesAutoresizingMaskIntoConstraints = false

        let content = window.contentView!
        content.addSubview(layout)
        NSLayoutConstraint.activate([
            layout.centerXAnchor.constraint(equalTo: content.centerXAnchor),
            layout.centerYAnchor.constraint(equalTo: content.centerYAnchor),
            note.widthAnchor.constraint(lessThanOrEqualTo: content.widthAnchor, constant: -32),
        ])
        window.center()
        window.makeKeyAndOrderFront(nil)
        NSApp.activate(ignoringOtherApps: true)
    }

    func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool {
        true
    }

    func windowShouldClose(_ sender: NSWindow) -> Bool {
        !busy
    }

    func applicationShouldTerminate(_ sender: NSApplication) -> NSApplication.TerminateReply {
        busy ? .terminateCancel : .terminateNow
    }

    @objc private func install() {
        if Bundle.main.bundleURL.pathComponents.contains("AppTranslocation") {
            show("请先移动 App", "请在 Finder 中将 App 移到「应用程序」，再从那里打开并安装。", error: true)
            return
        }
        run("manage.py", ["install"], success: "WPS 加载项已安装")
    }

    @objc private func uninstall() {
        run("manage.py", ["uninstall"], success: "WPS 加载项已卸载")
    }

    @objc private func convert() {
        let panel = NSOpenPanel()
        panel.title = "选择包含 LaTeX 公式的 .docx 文件"
        panel.allowsMultipleSelection = true
        panel.allowedContentTypes = [UTType(filenameExtension: "docx") ?? .data]
        panel.beginSheetModal(for: window) { [weak self] response in
            guard response == .OK else { return }
            self?.run("convert.py", panel.urls.map(\.path), success: "文件转换完成")
        }
    }

    private func run(_ script: String, _ arguments: [String], success: String) {
        guard !busy, let resources = Bundle.main.resourceURL else { return }
        let process = Process()
        let output = Pipe()
        process.executableURL = resources.appendingPathComponent("python/bin/python3")
        process.arguments = ["-BEs", resources.appendingPathComponent("src/\(script)").path] + arguments
        process.standardOutput = output
        process.standardError = output
        do {
            try process.run()
        } catch {
            show("无法启动转换服务", error.localizedDescription, error: true)
            return
        }
        busy = true
        buttons.forEach { $0.isEnabled = false }
        DispatchQueue.global(qos: .userInitiated).async { [self] in
            let data = output.fileHandleForReading.readDataToEndOfFile()
            process.waitUntilExit()
            let failed = process.terminationStatus != 0
            let detail = String(decoding: data, as: UTF8.self).trimmingCharacters(in: .whitespacesAndNewlines)
            DispatchQueue.main.async {
                self.busy = false
                self.buttons.forEach { $0.isEnabled = true }
                let title = failed
                    ? (script == "convert.py" ? "请查看转换结果" : "WPS 加载项配置失败")
                    : success
                self.show(title, detail, error: failed)
            }
        }
    }

    private func show(_ title: String, _ detail: String, error: Bool) {
        let alert = NSAlert()
        alert.messageText = title
        alert.alertStyle = error ? .warning : .informational
        if detail.count <= 450 {
            alert.informativeText = detail
        } else {
            let scroll = NSScrollView(frame: NSRect(x: 0, y: 0, width: 440, height: 180))
            let text = NSTextView(frame: scroll.bounds)
            text.string = detail
            text.isEditable = false
            text.isVerticallyResizable = true
            text.autoresizingMask = [.width]
            scroll.documentView = text
            scroll.hasVerticalScroller = true
            alert.accessoryView = scroll
        }
        alert.addButton(withTitle: "好")
        alert.beginSheetModal(for: window)
    }
}

let application = NSApplication.shared
let controller = Controller()
application.setActivationPolicy(.regular)
application.delegate = controller
application.run()
