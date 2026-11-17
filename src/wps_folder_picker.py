"""Request per-folder access with AppKit when macOS blocks the command script."""

from contextlib import contextmanager
from ctypes import CDLL, CFUNCTYPE, c_bool, c_char_p, c_long, c_void_p
from pathlib import Path


@contextmanager
def open_wps_folder(expected: Path):
    # Keep the picker and the subsequent file operations in the same process.
    CDLL('/System/Library/Frameworks/AppKit.framework/AppKit')
    objc = CDLL('/usr/lib/libobjc.A.dylib')
    get_class = objc.objc_getClass
    get_class.argtypes = [c_char_p]
    get_class.restype = c_void_p
    selector = objc.sel_registerName
    selector.argtypes = [c_char_p]
    selector.restype = c_void_p

    def message(result, *arguments):
        return CFUNCTYPE(result, c_void_p, c_void_p, *arguments)(('objc_msgSend', objc))

    get = message(c_void_p)
    get_text = message(c_char_p)
    get_number = message(c_long)
    get_flag = message(c_bool)
    send = message(None)
    set_object = message(None, c_void_p)
    set_flag = message(None, c_bool)
    set_number = message(c_bool, c_long)
    make_string = message(c_void_p, c_char_p)
    make_url = message(c_void_p, c_void_p)

    def ns_string(text: str):
        return make_string(get_class(b'NSString'), selector(b'stringWithUTF8String:'), text.encode('utf-8'))

    pool = get(get_class(b'NSAutoreleasePool'), selector(b'new'))
    try:
        app = get(get_class(b'NSApplication'), selector(b'sharedApplication'))
        set_number(app, selector(b'setActivationPolicy:'), 1)  # NSApplicationActivationPolicyAccessory
        send(app, selector(b'finishLaunching'))
        set_flag(app, selector(b'activateIgnoringOtherApps:'), True)
        panel = get(get_class(b'NSOpenPanel'), selector(b'openPanel'))
        set_flag(panel, selector(b'setCanChooseFiles:'), False)
        set_flag(panel, selector(b'setCanChooseDirectories:'), True)
        set_object(panel, selector(b'setTitle:'), ns_string('选择 WPS 加载项目录'))
        set_object(panel, selector(b'setMessage:'), ns_string(
            f'请选择已定位到的「{expected.name}」文件夹。安装器需要访问其中的加载项配置。'
        ))
        set_object(panel, selector(b'setPrompt:'), ns_string('选择此文件夹'))
        url = make_url(get_class(b'NSURL'), selector(b'fileURLWithPath:'), ns_string(str(expected)))
        set_object(panel, selector(b'setDirectoryURL:'), url)
        if get_number(panel, selector(b'runModal')) != 1:  # NSModalResponseOK
            raise RuntimeError('已取消 WPS 加载项配置。')

        selected = get(panel, selector(b'URL'))
        scoped = get_flag(selected, selector(b'startAccessingSecurityScopedResource'))
        try:
            selected_path = get_text(get(selected, selector(b'path')), selector(b'fileSystemRepresentation'))
            if selected_path is None or Path(selected_path.decode('utf-8')).resolve() != expected.resolve():
                raise RuntimeError(f'请选择 WPS 的「{expected.name}」文件夹后重试。')
            yield
        finally:
            if scoped:
                send(selected, selector(b'stopAccessingSecurityScopedResource'))
    finally:
        send(pool, selector(b'drain'))
