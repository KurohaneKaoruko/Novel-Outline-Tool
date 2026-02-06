 #![windows_subsystem = "windows"]
 
 use std::mem::{size_of, MaybeUninit};
use std::ffi::c_void;
 use std::path::PathBuf;
 use std::ptr::{null, null_mut};
 
 use windows_sys::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows_sys::Win32::Graphics::Gdi::{
    CreateFontW, CreateSolidBrush, DeleteObject, GetStockObject, SetBkColor, SetTextColor, CLIP_DEFAULT_PRECIS,
    CLEARTYPE_QUALITY, DEFAULT_CHARSET, DEFAULT_PITCH, FF_DONTCARE, FW_NORMAL, HBRUSH, HDC, OUT_DEFAULT_PRECIS,
    WHITE_BRUSH,
};
 use windows_sys::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
 use windows_sys::Win32::System::LibraryLoader::{GetModuleHandleW, LoadLibraryW};
 use windows_sys::Win32::UI::Controls::{
     InitCommonControlsEx, INITCOMMONCONTROLSEX, ICC_STANDARD_CLASSES, SB_SETTEXTW, STATUSCLASSNAMEW,
     TCM_GETCURSEL, TCM_INSERTITEMW, TCITEMW, TCN_SELCHANGE, TVGN_CARET, TVI_ROOT, TVIF_PARAM,
     TVIF_TEXT, TVINSERTSTRUCTW, TVITEMEXW, TVM_DELETEITEM, TVM_INSERTITEMW, TVM_SELECTITEM,
     TVHITTESTINFO, TVM_EDITLABELW, TVM_GETITEMW, TVM_GETNEXTITEM, TVM_HITTEST, TVM_SETBKCOLOR, TVM_SETLINECOLOR,
     TVM_SETTEXTCOLOR, TVN_BEGINDRAGW, TVN_ENDLABELEDITW, TVN_SELCHANGEDW, TVS_EDITLABELS, TVS_FULLROWSELECT,
     TVS_HASLINES, TVS_LINESATROOT, TVS_SHOWSELALWAYS, WC_TABCONTROLW, WC_TREEVIEWW, NMTVDISPINFOW, SB_SETPARTS,
 };
use windows::Win32::UI::Controls::RichEdit::{
    CHARFORMAT2W, CFE_AUTOCOLOR, CFE_BOLD, CFM_BOLD, CFM_COLOR, EM_SETCHARFORMAT, SCF_DEFAULT, SCF_SELECTION,
};
use windows_sys::Win32::UI::HiDpi::{GetDpiForWindow, SetProcessDpiAwarenessContext, DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2};
 use windows_sys::Win32::UI::WindowsAndMessaging::{
     CreateWindowExW, DefWindowProcW, DispatchMessageW, GetClientRect, GetMessageW, GetParent, LoadCursorW, LoadIconW, SetCursor,
     AppendMenuW, CreateMenu, CreatePopupMenu, DestroyWindow, DrawMenuBar, PostQuitMessage,
     RegisterClassExW, SendMessageW, SetMenu, SetWindowLongPtrW, ShowWindow, TranslateMessage,
     CW_USEDEFAULT, GWLP_USERDATA, HMENU, IDC_ARROW, IDC_SIZEWE, IDI_APPLICATION, MF_POPUP, MF_STRING, MSG, SW_SHOW, WM_COMMAND,
     WM_CREATE, WM_CTLCOLORDLG, WM_CTLCOLOREDIT, WM_CTLCOLORSTATIC, WM_DESTROY, WM_ERASEBKGND, WM_KEYDOWN, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE, WM_NCCREATE, WM_NOTIFY, WM_SETCURSOR, WM_SETFONT, WM_SIZE, WM_TIMER,
     WS_EX_CLIENTEDGE,
     ES_AUTOVSCROLL, ES_AUTOHSCROLL, ES_MULTILINE, EN_CHANGE, GetWindowTextLengthW,
     GetWindowTextW, KillTimer, SetTimer, SetWindowTextW, WS_HSCROLL, WS_VSCROLL,
     WNDCLASSEXW, WS_CHILD, WS_CLIPCHILDREN, WS_OVERLAPPEDWINDOW, WS_VISIBLE,
 };
 
 mod domain;
 mod storage;
mod update;
 
 use crate::domain::Project;
 use crate::storage::ProjectStore;
 
#[link(name = "user32")]
extern "system" {
    fn SetCapture(hWnd: isize) -> isize;
    fn ReleaseCapture() -> i32;
    fn GetCursorPos(lpPoint: *mut windows_sys::Win32::Foundation::POINT) -> i32;
    fn ScreenToClient(hWnd: isize, lpPoint: *mut windows_sys::Win32::Foundation::POINT) -> i32;
    fn FillRect(hDC: isize, lprc: *const windows_sys::Win32::Foundation::RECT, hbr: isize) -> i32;
    fn InvalidateRect(hWnd: isize, lpRect: *const windows_sys::Win32::Foundation::RECT, bErase: i32) -> i32;
    fn GetKeyState(nVirtKey: i32) -> i16;
    fn SetFocus(hWnd: isize) -> isize;
    fn GetFocus() -> isize;
}

#[link(name = "dwmapi")]
extern "system" {
    fn DwmSetWindowAttribute(hwnd: isize, dwAttribute: u32, pvAttribute: *const c_void, cbAttribute: u32) -> i32;
}

#[link(name = "uxtheme")]
extern "system" {
    fn SetWindowTheme(hwnd: isize, pszSubAppName: *const u16, pszSubIdList: *const u16) -> i32;
}

 const APP_CLASS: &str = "NovelOutlineToolMainWindow";
 const SPLITTER_CLASS: &str = "NovelOutlineToolSplitter";
 const APP_TITLE: &str = "Novel Outline Tool";
 const TREE_ID: isize = 1001;
 const EDIT_ID: isize = 1002;
 const STATUS_ID: isize = 1003;
 const TABS_ID: isize = 1004;
 const TIMER_AUTOSAVE: usize = 1;
 const TIMER_HIGHLIGHT: usize = 2;
 const TIMER_SEARCH: usize = 3;
 const SEARCH_ID: isize = 1005;
 const SPLITTER_ID: isize = 1006;
 
 const EM_GETSEL_MSG: u32 = 0x00B0;
 const EM_SETSEL_MSG: u32 = 0x00B1;
 const EM_GETLINE_MSG: u32 = 0x00C4;
 const EM_LINEFROMCHAR_MSG: u32 = 0x00C9;
 const EM_LINEINDEX_MSG: u32 = 0x00BB;
 const EM_LINELENGTH_MSG: u32 = 0x00C1;
 const EM_SETBKGNDCOLOR_MSG: u32 = 0x0443;
 const EM_SETMARGINS_MSG: u32 = 0x00D3;
 const EM_SETCUEBANNER_MSG: u32 = 0x1501;
 const DWMWA_USE_IMMERSIVE_DARK_MODE: u32 = 20;
 const EC_LEFTMARGIN: usize = 0x1;
 const EC_RIGHTMARGIN: usize = 0x2;
 
 const IDM_FILE_NEW: usize = 40001;
 const IDM_FILE_OPEN: usize = 40002;
 const IDM_FILE_SAVE: usize = 40003;
 const IDM_FILE_IMPORT: usize = 40005;
 const IDM_FILE_EXPORT: usize = 40006;
 const IDM_FILE_EXIT: usize = 40004;
 const IDM_ITEM_NEW: usize = 40101;
 const IDM_ITEM_RENAME: usize = 40102;
 const IDM_ITEM_DELETE: usize = 40103;
 const IDM_EDIT_UNDO: usize = 40201;
 const IDM_EDIT_REDO: usize = 40202;
 const IDM_VIEW_TOGGLE_THEME: usize = 40301;
 const IDM_HELP_CHECK_UPDATE: usize = 40401;
 
 fn wide(s: &str) -> Vec<u16> {
     let mut v: Vec<u16> = s.encode_utf16().collect();
     v.push(0);
     v
 }
 
 #[derive(Debug, Copy, Clone, PartialEq, Eq)]
 enum Section {
     Chapters,
     Characters,
     World,
     Timeline,
 }

#[derive(Debug, Clone)]
enum Command {
    CreateFile { path: PathBuf, contents: String },
    DeleteFile { path: PathBuf, contents: String },
    RenameFile { from: PathBuf, to: PathBuf },
    ReorderChapters { pairs: Vec<(PathBuf, PathBuf)> },
}

impl Command {
    fn apply(&self) -> Result<(), String> {
        match self {
            Command::CreateFile { path, contents } => crate::storage::write_text_atomic(path, contents).map_err(|e| e.to_string()),
            Command::DeleteFile { path, .. } => std::fs::remove_file(path).map_err(|e| e.to_string()),
            Command::RenameFile { from, to } => std::fs::rename(from, to).map_err(|e| e.to_string()),
            Command::ReorderChapters { pairs } => apply_rename_pairs(pairs).map_err(|e| e.to_string()),
        }
    }

    fn undo(&self) -> Result<(), String> {
        match self {
            Command::CreateFile { path, .. } => std::fs::remove_file(path).map_err(|e| e.to_string()),
            Command::DeleteFile { path, contents } => crate::storage::write_text_atomic(path, contents).map_err(|e| e.to_string()),
            Command::RenameFile { from, to } => std::fs::rename(to, from).map_err(|e| e.to_string()),
            Command::ReorderChapters { pairs } => {
                let reversed: Vec<(PathBuf, PathBuf)> = pairs.iter().map(|(a, b)| (b.clone(), a.clone())).collect();
                apply_rename_pairs(&reversed).map_err(|e| e.to_string())
            }
        }
    }
}
 
 struct AppState {
     hwnd_status: HWND,
     hwnd_tabs: HWND,
     hwnd_search: HWND,
     hwnd_tree: HWND,
     hwnd_edit: HWND,
    hwnd_splitter: HWND,
    hfont_ui: isize,
    left_pane_ratio: f32,
    hbr_light: HBRUSH,
    hbr_dark: HBRUSH,
     project: Option<Project>,
     current_section: Section,
     item_paths: Vec<PathBuf>,
     current_doc_path: Option<PathBuf>,
     current_doc_dirty: bool,
     filter_text: String,
     last_highlight_line: i32,
    undo_stack: Vec<Command>,
    redo_stack: Vec<Command>,
    dragging: bool,
    drag_src_idx: usize,
    last_backup_unix: u64,
    dark_mode: bool,
 }
 
 fn main() {
     unsafe {
         let _ = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
         let mut icc = INITCOMMONCONTROLSEX {
             dwSize: size_of::<INITCOMMONCONTROLSEX>() as u32,
             dwICC: ICC_STANDARD_CLASSES,
         };
         InitCommonControlsEx(&mut icc);
     }
 
     if let Err(message) = run() {
         unsafe {
             let title = wide("Fatal Error");
             let msg = wide(&message);
             windows_sys::Win32::UI::WindowsAndMessaging::MessageBoxW(
                 0,
                 msg.as_ptr(),
                 title.as_ptr(),
                 windows_sys::Win32::UI::WindowsAndMessaging::MB_ICONERROR,
             );
         }
     }
 }
 
 fn run() -> Result<(), String> {
     unsafe {
         let hr = CoInitializeEx(null_mut(), COINIT_APARTMENTTHREADED as u32);
         if hr < 0 {
             return Err("CoInitializeEx failed".to_string());
         }
 
         let hinstance = GetModuleHandleW(null());
         if hinstance == 0 {
             CoUninitialize();
             return Err("GetModuleHandleW failed".to_string());
         }
 
         let class_name = wide(APP_CLASS);
         let cursor = LoadCursorW(0, IDC_ARROW);
        let icon = LoadIconW(0, IDI_APPLICATION);
         let hbr_background: HBRUSH = GetStockObject(WHITE_BRUSH as i32) as HBRUSH;
 
         let wc = WNDCLASSEXW {
             cbSize: size_of::<WNDCLASSEXW>() as u32,
             style: 0,
             lpfnWndProc: Some(window_proc),
             cbClsExtra: 0,
             cbWndExtra: 0,
             hInstance: hinstance,
            hIcon: icon,
             hCursor: cursor,
             hbrBackground: hbr_background,
             lpszMenuName: null(),
             lpszClassName: class_name.as_ptr(),
            hIconSm: icon,
         };
 
         if RegisterClassExW(&wc) == 0 {
             CoUninitialize();
             return Err("RegisterClassExW failed".to_string());
         }
 
         let splitter_class = wide(SPLITTER_CLASS);
         let wc_splitter = WNDCLASSEXW {
             cbSize: size_of::<WNDCLASSEXW>() as u32,
             style: 0,
             lpfnWndProc: Some(splitter_proc),
             cbClsExtra: 0,
             cbWndExtra: 0,
             hInstance: hinstance,
             hIcon: 0,
             hCursor: LoadCursorW(0, IDC_SIZEWE),
             hbrBackground: hbr_background,
             lpszMenuName: null(),
             lpszClassName: splitter_class.as_ptr(),
             hIconSm: 0,
         };
         if RegisterClassExW(&wc_splitter) == 0 {
             CoUninitialize();
             return Err("RegisterClassExW splitter failed".to_string());
         }
 
         let title = wide(APP_TITLE);
         let hwnd = CreateWindowExW(
             0,
             class_name.as_ptr(),
             title.as_ptr(),
             WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
             CW_USEDEFAULT,
             CW_USEDEFAULT,
             1200,
             800,
             0,
             0,
             hinstance,
             null_mut(),
         );
 
         if hwnd == 0 {
             CoUninitialize();
             return Err("CreateWindowExW failed".to_string());
         }
 
         ShowWindow(hwnd, SW_SHOW);
 
         let mut msg = MaybeUninit::<MSG>::uninit();
         loop {
             let ret = GetMessageW(msg.as_mut_ptr(), 0, 0, 0);
             if ret == 0 {
                 break;
             }
             if ret == -1 {
                 CoUninitialize();
                 return Err("GetMessageW failed".to_string());
             }
             let msg = msg.assume_init();
             TranslateMessage(&msg);
             DispatchMessageW(&msg);
         }
 
         CoUninitialize();
         Ok(())
     }
 }
 
 unsafe extern "system" fn window_proc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
     match msg {
         WM_NCCREATE => {
             let state = Box::new(AppState {
                 hwnd_status: 0,
                 hwnd_tabs: 0,
                 hwnd_search: 0,
                 hwnd_tree: 0,
                 hwnd_edit: 0,
                hwnd_splitter: 0,
                hfont_ui: 0,
                left_pane_ratio: 0.28,
                hbr_light: 0,
                hbr_dark: 0,
                 project: None,
                 current_section: Section::Chapters,
                 item_paths: Vec::new(),
                 current_doc_path: None,
                 current_doc_dirty: false,
                 filter_text: String::new(),
                 last_highlight_line: -1,
                 undo_stack: Vec::new(),
                 redo_stack: Vec::new(),
                 dragging: false,
                 drag_src_idx: 0,
                 last_backup_unix: 0,
                 dark_mode: false,
             });
             SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
             DefWindowProcW(hwnd, msg, wparam, lparam)
         }
         WM_CREATE => {
             let state = state(hwnd);
 
             let menu = CreateMenu();
             let file_menu = CreatePopupMenu();
             AppendMenuW(file_menu, MF_STRING, IDM_FILE_NEW, wide("新建/初始化项目...").as_ptr());
             AppendMenuW(file_menu, MF_STRING, IDM_FILE_OPEN, wide("打开项目文件夹...").as_ptr());
             AppendMenuW(file_menu, MF_STRING, IDM_FILE_SAVE, wide("保存").as_ptr());
             AppendMenuW(file_menu, MF_STRING, IDM_FILE_IMPORT, wide("从文件夹导入...").as_ptr());
             AppendMenuW(file_menu, MF_STRING, IDM_FILE_EXPORT, wide("导出为文件夹...").as_ptr());
             AppendMenuW(file_menu, MF_STRING, IDM_ITEM_NEW, wide("新建当前条目").as_ptr());
             AppendMenuW(file_menu, MF_STRING, IDM_ITEM_RENAME, wide("重命名当前条目").as_ptr());
             AppendMenuW(file_menu, MF_STRING, IDM_ITEM_DELETE, wide("删除当前条目").as_ptr());
             AppendMenuW(file_menu, MF_STRING, IDM_FILE_EXIT, wide("退出").as_ptr());
             AppendMenuW(menu, MF_POPUP, file_menu as usize, wide("文件").as_ptr());
 
             let edit_menu = CreatePopupMenu();
             AppendMenuW(edit_menu, MF_STRING, IDM_EDIT_UNDO, wide("撤销结构操作").as_ptr());
             AppendMenuW(edit_menu, MF_STRING, IDM_EDIT_REDO, wide("重做结构操作").as_ptr());
             AppendMenuW(menu, MF_POPUP, edit_menu as usize, wide("编辑").as_ptr());
 
             let view_menu = CreatePopupMenu();
             AppendMenuW(view_menu, MF_STRING, IDM_VIEW_TOGGLE_THEME, wide("深色/浅色主题").as_ptr());
             AppendMenuW(menu, MF_POPUP, view_menu as usize, wide("视图").as_ptr());
 
             let help_menu = CreatePopupMenu();
             AppendMenuW(help_menu, MF_STRING, IDM_HELP_CHECK_UPDATE, wide("检查更新...").as_ptr());
             AppendMenuW(menu, MF_POPUP, help_menu as usize, wide("帮助").as_ptr());
 
             SetMenu(hwnd, menu);
             DrawMenuBar(hwnd);
 
             let _ = LoadLibraryW(wide("Msftedit.dll").as_ptr());

             state.hwnd_status = CreateWindowExW(
                 0,
                 STATUSCLASSNAMEW,
                 null(),
                 WS_CHILD | WS_VISIBLE,
                 0,
                 0,
                 0,
                 0,
                 hwnd,
                 STATUS_ID as HMENU,
                 GetModuleHandleW(null()),
                 null_mut(),
             );
 
             state.hwnd_tabs = CreateWindowExW(
                 0,
                 WC_TABCONTROLW,
                 null(),
                 WS_CHILD | WS_VISIBLE,
                 0,
                 0,
                 0,
                 0,
                 hwnd,
                 TABS_ID as HMENU,
                 GetModuleHandleW(null()),
                 null_mut(),
             );
             add_tabs(state.hwnd_tabs);

             state.hwnd_search = CreateWindowExW(
                 WS_EX_CLIENTEDGE,
                 wide("EDIT").as_ptr(),
                 null(),
                 WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL as u32,
                 0,
                 0,
                 0,
                 0,
                 hwnd,
                 SEARCH_ID as HMENU,
                 GetModuleHandleW(null()),
                 null_mut(),
             );
             let dpi = GetDpiForWindow(hwnd);
             let margin = scale_px(dpi, 6);
             SendMessageW(
                 state.hwnd_search,
                 EM_SETMARGINS_MSG,
                 EC_LEFTMARGIN | EC_RIGHTMARGIN,
                 make_lparam_u16(margin, margin),
             );
             let cue = wide("搜索当前模块...");
             SendMessageW(state.hwnd_search, EM_SETCUEBANNER_MSG, 0, cue.as_ptr() as LPARAM);

             state.hwnd_tree = CreateWindowExW(
                 WS_EX_CLIENTEDGE,
                 WC_TREEVIEWW,
                 null(),
                 WS_CHILD
                     | WS_VISIBLE
                     | TVS_EDITLABELS as u32
                     | TVS_FULLROWSELECT as u32
                     | TVS_SHOWSELALWAYS as u32
                     | TVS_HASLINES as u32
                     | TVS_LINESATROOT as u32,
                 0,
                 0,
                 0,
                 0,
                 hwnd,
                 TREE_ID as HMENU,
                 GetModuleHandleW(null()),
                 null_mut(),
             );

             state.hwnd_splitter = CreateWindowExW(
                 0,
                 wide(SPLITTER_CLASS).as_ptr(),
                 null(),
                 WS_CHILD | WS_VISIBLE,
                 0,
                 0,
                 0,
                 0,
                 hwnd,
                 SPLITTER_ID as HMENU,
                 GetModuleHandleW(null()),
                 null_mut(),
             );

             state.hwnd_edit = CreateWindowExW(
                 WS_EX_CLIENTEDGE,
                 wide("RICHEDIT50W").as_ptr(),
                 null(),
                 WS_CHILD
                     | WS_VISIBLE
                     | WS_VSCROLL
                     | WS_HSCROLL
                     | ES_MULTILINE as u32
                     | ES_AUTOVSCROLL as u32
                     | ES_AUTOHSCROLL as u32,
                 0,
                 0,
                 0,
                 0,
                 hwnd,
                 EDIT_ID as HMENU,
                 GetModuleHandleW(null()),
                 null_mut(),
             );
 
             let dpi = GetDpiForWindow(hwnd);
             let margin = scale_px(dpi, 12);
             SendMessageW(
                 state.hwnd_edit,
                 EM_SETMARGINS_MSG,
                 EC_LEFTMARGIN | EC_RIGHTMARGIN,
                 make_lparam_u16(margin, margin),
             );
 
             state.hbr_light = CreateSolidBrush(0x00FFFFFF);
             state.hbr_dark = CreateSolidBrush(0x00202020);
             apply_theme(hwnd, state.dark_mode);
 
             state.hfont_ui = create_ui_font(hwnd);
             apply_ui_font(hwnd, state.hfont_ui);
             layout(hwnd);
 
             set_status_text(hwnd, "就绪");
             0
         }
         WM_COMMAND => {
             if lparam != 0 {
                 let hwnd_from = lparam as HWND;
                 let code = ((wparam as u32 >> 16) & 0xffff) as u16;
                if code as u32 == EN_CHANGE {
                    let state = state(hwnd);
                    if hwnd_from == state.hwnd_edit {
                        state.current_doc_dirty = true;
                        let _ = SetTimer(hwnd, TIMER_HIGHLIGHT, 250, None);
                        return 0;
                    }
                    if hwnd_from == state.hwnd_search {
                        let _ = SetTimer(hwnd, TIMER_SEARCH, 150, None);
                        return 0;
                    }
                }
                 return DefWindowProcW(hwnd, msg, wparam, lparam);
             }

             let id = (wparam as u32 & 0xffff) as usize;
             match id {
                 IDM_FILE_NEW | IDM_FILE_OPEN => {
                     if let Some(root) = pick_folder(hwnd) {
                         match ProjectStore::open_or_init(root) {
                             Ok(project) => {
                                 let state = state(hwnd);
                                 state.project = Some(project);
                                 state.dark_mode = state
                                     .project
                                     .as_ref()
                                     .and_then(|p| p.meta.theme.as_deref())
                                     .map(|t| t.eq_ignore_ascii_case("dark"))
                                     .unwrap_or(false);
                                 apply_theme(hwnd, state.dark_mode);
                                 state.left_pane_ratio = state
                                     .project
                                     .as_ref()
                                     .and_then(|p| p.meta.left_pane_ratio)
                                     .unwrap_or(0.28);
                                 state.current_section = Section::Chapters;
                                 state.item_paths.clear();
                                 state.current_doc_path = None;
                                 state.current_doc_dirty = false;
                                 state.filter_text.clear();
                                 state.last_highlight_line = -1;
                                 state.undo_stack.clear();
                                 state.redo_stack.clear();
                                 state.dragging = false;
                                 state.last_backup_unix = 0;
                                 SetWindowTextW(state.hwnd_search, wide("").as_ptr());
                                 SendMessageW(state.hwnd_tree, TVM_DELETEITEM, 0, 0);
                                 SendMessageW(state.hwnd_tabs, windows_sys::Win32::UI::Controls::TCM_SETCURSEL, 0, 0);
                                 reload_items(hwnd);
                                 let _ = SetTimer(hwnd, TIMER_AUTOSAVE, 30_000, None);
                                 if let Some(project) = &state.project {
                                     set_status_text(hwnd, &format!("已打开: {}", project.root.display()));
                                     set_status_part(hwnd, 1, "");
                                     set_status_part(hwnd, 2, "已保存");
                                 }
                             }
                             Err(e) => show_error(hwnd, &e),
                         }
                     }
                     0
                 }
                IDM_ITEM_NEW => {
                    if let Err(e) = create_new_item(hwnd) {
                        show_error(hwnd, &e);
                    }
                    0
                }
                IDM_ITEM_RENAME => {
                    begin_rename_selected(hwnd);
                    0
                }
                IDM_ITEM_DELETE => {
                    if let Err(e) = delete_selected_item(hwnd) {
                        show_error(hwnd, &e);
                    }
                    0
                }
                 IDM_FILE_SAVE => {
                     save_current_if_dirty(hwnd);
                     let state = state(hwnd);
                     if let Some(project) = &state.project {
                         if let Err(e) = ProjectStore::save_project_meta(project) {
                             show_error(hwnd, &e);
                         } else {
                             set_status_text(hwnd, "已保存");
                         }
                     }
                     0
                 }
                IDM_FILE_IMPORT => {
                    if let Err(e) = import_project(hwnd) {
                        show_error(hwnd, &e);
                    }
                    0
                }
                IDM_FILE_EXPORT => {
                    if let Err(e) = export_project(hwnd) {
                        show_error(hwnd, &e);
                    }
                    0
                }
                IDM_EDIT_UNDO => {
                    if let Err(e) = do_undo(hwnd) {
                        show_error(hwnd, &e);
                    }
                    0
                }
                IDM_EDIT_REDO => {
                    if let Err(e) = do_redo(hwnd) {
                        show_error(hwnd, &e);
                    }
                    0
                }
                IDM_VIEW_TOGGLE_THEME => {
                    let state = state(hwnd);
                    state.dark_mode = !state.dark_mode;
                    apply_theme(hwnd, state.dark_mode);
                    if let Some(project) = &mut state.project {
                        project.meta.theme = Some(if state.dark_mode { "dark".to_string() } else { "light".to_string() });
                        let _ = ProjectStore::save_project_meta(project);
                    }
                    0
                }
                IDM_HELP_CHECK_UPDATE => {
                    if let Err(e) = check_updates(hwnd) {
                        show_error(hwnd, &e);
                    }
                    0
                }
                 IDM_FILE_EXIT => {
                     DestroyWindow(hwnd);
                     0
                 }
                 _ => DefWindowProcW(hwnd, msg, wparam, lparam),
             }
         }
        WM_CTLCOLORDLG | WM_CTLCOLOREDIT | WM_CTLCOLORSTATIC => {
            let state = state(hwnd);
            let hdc = wparam as HDC;
            let (bg, fg, hbr) = if state.dark_mode {
                (0x00202020u32, 0x00E0E0E0u32, state.hbr_dark)
            } else {
                (0x00FFFFFFu32, 0x00000000u32, state.hbr_light)
            };
            SetBkColor(hdc, bg);
            SetTextColor(hdc, fg);
            hbr as LRESULT
        }
        WM_ERASEBKGND => {
            let state = state(hwnd);
            let hdc = wparam as HDC;
            let mut rc = windows_sys::Win32::Foundation::RECT {
                left: 0,
                top: 0,
                right: 0,
                bottom: 0,
            };
            GetClientRect(hwnd, &mut rc);
            let hbr = if state.dark_mode { state.hbr_dark } else { state.hbr_light };
            FillRect(hdc, &rc, hbr);
            1
        }
        WM_KEYDOWN => {
            let ctrl = (GetKeyState(0x11) as i32) < 0;
            let shift = (GetKeyState(0x10) as i32) < 0;
            let alt = (GetKeyState(0x12) as i32) < 0;
            let vk = wparam as u32;
            let focus = GetFocus();
            let state = state(hwnd);

            if ctrl && !alt && !shift && vk == ('O' as u32) {
                SendMessageW(hwnd, WM_COMMAND, IDM_FILE_OPEN, 0);
                return 0;
            }
            if ctrl && !alt && !shift && vk == ('S' as u32) {
                SendMessageW(hwnd, WM_COMMAND, IDM_FILE_SAVE, 0);
                return 0;
            }
            if ctrl && !alt && !shift && vk == ('N' as u32) {
                SendMessageW(hwnd, WM_COMMAND, IDM_FILE_NEW, 0);
                return 0;
            }
            if ctrl && !alt && !shift && vk == ('F' as u32) {
                if state.hwnd_search != 0 {
                    SetFocus(state.hwnd_search);
                }
                return 0;
            }
            if ctrl && alt && vk == ('Z' as u32) {
                SendMessageW(hwnd, WM_COMMAND, IDM_EDIT_UNDO, 0);
                return 0;
            }
            if ctrl && alt && vk == ('Y' as u32) {
                SendMessageW(hwnd, WM_COMMAND, IDM_EDIT_REDO, 0);
                return 0;
            }
            if vk == 0x71 && focus != state.hwnd_edit {
                SendMessageW(hwnd, WM_COMMAND, IDM_ITEM_RENAME, 0);
                return 0;
            }
            if vk == 0x2E && focus != state.hwnd_edit {
                SendMessageW(hwnd, WM_COMMAND, IDM_ITEM_DELETE, 0);
                return 0;
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
         WM_NOTIFY => {
             let state = state(hwnd);
             if lparam == 0 {
                 return 0;
             }
             let hdr = &*(lparam as *const windows_sys::Win32::UI::Controls::NMHDR);
             if hdr.hwndFrom == state.hwnd_tabs && hdr.code as u32 == TCN_SELCHANGE {
                 save_current_if_dirty(hwnd);
                 let idx = SendMessageW(state.hwnd_tabs, TCM_GETCURSEL, 0, 0) as i32;
                 state.current_section = match idx {
                     1 => Section::Characters,
                     2 => Section::World,
                     3 => Section::Timeline,
                     _ => Section::Chapters,
                 };
                 SendMessageW(state.hwnd_tree, TVM_DELETEITEM, 0, 0);
                 reload_items(hwnd);
                 return 0;
             }
             if hdr.hwndFrom == state.hwnd_tree && hdr.code as u32 == TVN_BEGINDRAGW {
                 let nmtv = &*(lparam as *const windows_sys::Win32::UI::Controls::NMTREEVIEWW);
                 let idx = nmtv.itemNew.lParam as isize;
                 if idx >= 0 {
                     state.dragging = true;
                     state.drag_src_idx = idx as usize;
                     SetCapture(hwnd);
                 }
                 return 0;
             }
             if hdr.hwndFrom == state.hwnd_tree && hdr.code as u32 == TVN_ENDLABELEDITW {
                 let disp = &*(lparam as *const NMTVDISPINFOW);
                 let idx = disp.item.lParam as isize;
                 if idx < 0 {
                     return 0;
                 }
                 if disp.item.pszText.is_null() {
                     return 0;
                 }
                 let new_name = wide_ptr_to_string(disp.item.pszText);
                 match commit_rename(hwnd, idx as usize, &new_name) {
                     Ok(true) => return 1,
                     Ok(false) => return 0,
                     Err(e) => {
                         show_error(hwnd, &e);
                         return 0;
                     }
                 }
             }
             if hdr.hwndFrom == state.hwnd_tree && hdr.code as u32 == TVN_SELCHANGEDW {
                 let nmtv = &*(lparam as *const windows_sys::Win32::UI::Controls::NMTREEVIEWW);
                 let idx = nmtv.itemNew.lParam as isize;
                 if idx >= 0 {
                     open_item_by_index(hwnd, idx as usize);
                 }
                 return 0;
             }
             DefWindowProcW(hwnd, msg, wparam, lparam)
         }
         WM_MOUSEMOVE => {
             let state = state(hwnd);
             if !state.dragging {
                 return DefWindowProcW(hwnd, msg, wparam, lparam);
             }
             drag_update_hover(hwnd);
             0
         }
         WM_LBUTTONUP => {
             let state = state(hwnd);
             if !state.dragging {
                 return DefWindowProcW(hwnd, msg, wparam, lparam);
             }
             ReleaseCapture();
             state.dragging = false;
             if let Err(e) = drag_commit_drop(hwnd, state.drag_src_idx) {
                 show_error(hwnd, &e);
             }
             0
         }
         WM_SIZE => {
             layout(hwnd);
             0
         }
         WM_TIMER => {
             if wparam == TIMER_AUTOSAVE {
                 save_current_if_dirty(hwnd);
             }
            if wparam == TIMER_SEARCH {
                KillTimer(hwnd, TIMER_SEARCH);
                let s = get_text(state(hwnd).hwnd_search);
                state(hwnd).filter_text = s;
                SendMessageW(state(hwnd).hwnd_tree, TVM_DELETEITEM, 0, 0);
                reload_items(hwnd);
            }
            if wparam == TIMER_HIGHLIGHT {
                KillTimer(hwnd, TIMER_HIGHLIGHT);
                highlight_current_line(hwnd);
            }
             0
         }
         WM_DESTROY => {
            let _ = KillTimer(hwnd, TIMER_AUTOSAVE);
            let _ = KillTimer(hwnd, TIMER_SEARCH);
            let _ = KillTimer(hwnd, TIMER_HIGHLIGHT);
             let ptr = windows_sys::Win32::UI::WindowsAndMessaging::GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AppState;
             if !ptr.is_null() {
                 if (*ptr).hfont_ui != 0 {
                     let _ = DeleteObject((*ptr).hfont_ui);
                     (*ptr).hfont_ui = 0;
                 }
                 if (*ptr).hbr_light != 0 {
                     let _ = DeleteObject((*ptr).hbr_light);
                     (*ptr).hbr_light = 0;
                 }
                 if (*ptr).hbr_dark != 0 {
                     let _ = DeleteObject((*ptr).hbr_dark);
                     (*ptr).hbr_dark = 0;
                 }
                 drop(Box::from_raw(ptr));
             }
             PostQuitMessage(0);
             0
         }
         _ => DefWindowProcW(hwnd, msg, wparam, lparam),
     }
 }
 
 unsafe fn state(hwnd: HWND) -> &'static mut AppState {
     let ptr = windows_sys::Win32::UI::WindowsAndMessaging::GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AppState;
     &mut *ptr
 }
 
fn scale_px(dpi: u32, px: i32) -> i32 {
    ((px as i64 * dpi as i64) / 96) as i32
}

fn make_lparam_u16(lo: i32, hi: i32) -> LPARAM {
    let lo = (lo as u32) & 0xffff;
    let hi = (hi as u32) & 0xffff;
    ((lo | (hi << 16)) as isize) as LPARAM
}

unsafe fn create_ui_font(hwnd: HWND) -> isize {
    let dpi = GetDpiForWindow(hwnd) as i32;
    let height = -((9 * dpi + 36) / 72);
    let face = wide("Segoe UI");
    CreateFontW(
        height,
        0,
        0,
        0,
        FW_NORMAL as i32,
        0,
        0,
        0,
        DEFAULT_CHARSET as u32,
        OUT_DEFAULT_PRECIS as u32,
        CLIP_DEFAULT_PRECIS as u32,
        CLEARTYPE_QUALITY as u32,
        (DEFAULT_PITCH | FF_DONTCARE) as u32,
        face.as_ptr(),
    ) as isize
}

unsafe fn apply_ui_font(hwnd: HWND, hfont: isize) {
    if hfont == 0 {
        return;
    }
    let state = state(hwnd);
    let ctrls = [state.hwnd_status, state.hwnd_tabs, state.hwnd_search, state.hwnd_tree, state.hwnd_edit];
    for c in ctrls {
        if c != 0 {
            SendMessageW(c, WM_SETFONT, hfont as usize, 1);
        }
    }
}

unsafe extern "system" fn splitter_proc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_LBUTTONDOWN => {
            SetCapture(hwnd);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 1);
            0
        }
        WM_MOUSEMOVE => {
            if windows_sys::Win32::UI::WindowsAndMessaging::GetWindowLongPtrW(hwnd, GWLP_USERDATA) == 0 {
                return DefWindowProcW(hwnd, msg, wparam, lparam);
            }
            let parent = GetParent(hwnd);
            if parent != 0 {
                let mut pt = windows_sys::Win32::Foundation::POINT { x: 0, y: 0 };
                if GetCursorPos(&mut pt) != 0 {
                    ScreenToClient(parent, &mut pt);
                    set_left_pane_from_px(parent, pt.x);
                    layout(parent);
                }
            }
            0
        }
        WM_LBUTTONUP => {
            if windows_sys::Win32::UI::WindowsAndMessaging::GetWindowLongPtrW(hwnd, GWLP_USERDATA) != 0 {
                ReleaseCapture();
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                let parent = GetParent(hwnd);
                if parent != 0 {
                    persist_left_pane_ratio(parent);
                }
            }
            0
        }
        WM_SETCURSOR => {
            SetCursor(LoadCursorW(0, IDC_SIZEWE));
            1
        }
        WM_ERASEBKGND => {
            let parent = GetParent(hwnd);
            if parent != 0 {
                let st = state(parent);
                let hdc = wparam as HDC;
                let mut rc = windows_sys::Win32::Foundation::RECT {
                    left: 0,
                    top: 0,
                    right: 0,
                    bottom: 0,
                };
                GetClientRect(hwnd, &mut rc);
                let hbr = if st.dark_mode { st.hbr_dark } else { st.hbr_light };
                FillRect(hdc, &rc, hbr);
                return 1;
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn set_left_pane_from_px(hwnd: HWND, left_px: i32) {
    let mut rc = windows_sys::Win32::Foundation::RECT {
        left: 0,
        top: 0,
        right: 0,
        bottom: 0,
    };
    GetClientRect(hwnd, &mut rc);
    let width = (rc.right - rc.left).max(1);

    let dpi = GetDpiForWindow(hwnd);
    let min_left = scale_px(dpi, 240);
    let min_right = scale_px(dpi, 420);
    let max_left = (width - min_right).max(min_left);
    let left = left_px.clamp(min_left, max_left);
    state(hwnd).left_pane_ratio = (left as f32 / width as f32).clamp(0.18, 0.72);
}

unsafe fn persist_left_pane_ratio(hwnd: HWND) {
    let state = state(hwnd);
    if let Some(project) = &mut state.project {
        project.meta.left_pane_ratio = Some(state.left_pane_ratio);
        let _ = ProjectStore::save_project_meta(project);
    }
}

 unsafe fn layout(hwnd: HWND) {
     let state = state(hwnd);
     let mut rc = windows_sys::Win32::Foundation::RECT {
         left: 0,
         top: 0,
         right: 0,
         bottom: 0,
     };
     GetClientRect(hwnd, &mut rc);
 
     if state.hwnd_status != 0 {
         SendMessageW(state.hwnd_status, windows_sys::Win32::UI::WindowsAndMessaging::WM_SIZE, 0, 0);
     }
 
     let mut rc_status = windows_sys::Win32::Foundation::RECT {
         left: 0,
         top: 0,
         right: 0,
         bottom: 0,
     };
     if state.hwnd_status != 0 {
         windows_sys::Win32::UI::WindowsAndMessaging::GetWindowRect(state.hwnd_status, &mut rc_status);
     }
 
     let status_height = (rc_status.bottom - rc_status.top).max(0);
     let width = (rc.right - rc.left).max(0);
     let height = (rc.bottom - rc.top - status_height).max(0);
 
     if state.hwnd_status != 0 {
         let parts = [(width * 55) / 100, (width * 85) / 100, -1];
         SendMessageW(state.hwnd_status, SB_SETPARTS, parts.len(), parts.as_ptr() as LPARAM);
     }
 
    let dpi = GetDpiForWindow(hwnd);
    let padding = scale_px(dpi, 8);
    let gap = scale_px(dpi, 6);

    let splitter_w = scale_px(dpi, 6);
    let min_left = scale_px(dpi, 240);
    let min_right = scale_px(dpi, 420);
    let mut left_width = (width as f32 * state.left_pane_ratio) as i32;
    let max_left = (width - min_right - splitter_w).max(min_left);
    left_width = left_width.clamp(min_left, max_left);
    let right_width = width - left_width - splitter_w;
    let tabs_height = scale_px(dpi, 32);
    let search_height = scale_px(dpi, 28);
 
     if state.hwnd_tabs != 0 {
         windows_sys::Win32::UI::WindowsAndMessaging::MoveWindow(
             state.hwnd_tabs,
            padding,
            padding,
            (left_width - padding * 2).max(0),
             tabs_height,
             1,
         );
     }

    if state.hwnd_search != 0 {
        windows_sys::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.hwnd_search,
            padding,
            padding + tabs_height + gap,
            (left_width - padding * 2).max(0),
            search_height,
            1,
        );
    }

     if state.hwnd_tree != 0 {
        let tree_y = padding + tabs_height + gap + search_height + gap;
         windows_sys::Win32::UI::WindowsAndMessaging::MoveWindow(
             state.hwnd_tree,
            padding,
            tree_y,
            (left_width - padding * 2).max(0),
            (height - tree_y - padding).max(0),
             1,
         );
     }
 
     if state.hwnd_edit != 0 {
         windows_sys::Win32::UI::WindowsAndMessaging::MoveWindow(
             state.hwnd_edit,
            left_width + splitter_w + padding,
            padding,
            (right_width - padding * 2).max(0),
            (height - padding * 2).max(0),
             1,
         );
     }

    if state.hwnd_splitter != 0 {
        windows_sys::Win32::UI::WindowsAndMessaging::MoveWindow(
            state.hwnd_splitter,
            left_width,
            padding,
            splitter_w.max(1),
            (height - padding * 2).max(0),
            1,
        );
    }
 }

 unsafe fn reload_items(hwnd: HWND) {
     let state = state(hwnd);
     let Some(project) = &state.project else { return };

     let (dir_name, root_label) = match state.current_section {
         Section::Chapters => ("chapters", "章节"),
         Section::Characters => ("characters", "角色"),
         Section::World => ("world", "世界观"),
         Section::Timeline => ("timeline", "时间线"),
     };
 
     let dir = project.root.join(dir_name);
     let paths = crate::storage::list_markdown_files(&dir).unwrap_or_default();
    let filter = state.filter_text.trim().to_lowercase();
    if filter.is_empty() {
        state.item_paths = paths;
    } else {
        state.item_paths = paths
            .into_iter()
            .filter(|p| {
                p.file_stem()
                    .and_then(|s| s.to_str())
                    .map(|s| s.to_lowercase().contains(&filter))
                    .unwrap_or(false)
            })
            .collect();
    }

     let mut root_text = wide(root_label);
     let root_item = TVITEMEXW {
         mask: TVIF_TEXT as u32,
         hItem: 0,
         state: 0,
         stateMask: 0,
         pszText: root_text.as_mut_ptr(),
         cchTextMax: 0,
         iImage: 0,
         iSelectedImage: 0,
         cChildren: 0,
         lParam: -1,
         iIntegral: 0,
         uStateEx: 0,
         hwnd: 0,
         iExpandedImage: 0,
         iReserved: 0,
     };
     let mut ins = TVINSERTSTRUCTW {
         hParent: TVI_ROOT as isize,
         hInsertAfter: windows_sys::Win32::UI::Controls::TVI_LAST as isize,
         Anonymous: windows_sys::Win32::UI::Controls::TVINSERTSTRUCTW_0 { itemex: root_item },
     };
     let root = SendMessageW(state.hwnd_tree, TVM_INSERTITEMW, 0, &mut ins as *mut _ as LPARAM) as isize;

     let mut first_child: isize = 0;
     for (i, path) in state.item_paths.iter().enumerate() {
         let name = path
             .file_stem()
             .and_then(|s| s.to_str())
             .unwrap_or("chapter")
             .to_string();
         let mut text = wide(&name);
         let item = TVITEMEXW {
             mask: (TVIF_TEXT | TVIF_PARAM) as u32,
             hItem: 0,
             state: 0,
             stateMask: 0,
             pszText: text.as_mut_ptr(),
             cchTextMax: 0,
             iImage: 0,
             iSelectedImage: 0,
             cChildren: 0,
             lParam: i as isize,
             iIntegral: 0,
             uStateEx: 0,
             hwnd: 0,
             iExpandedImage: 0,
             iReserved: 0,
         };
         let mut ins = TVINSERTSTRUCTW {
             hParent: root,
             hInsertAfter: windows_sys::Win32::UI::Controls::TVI_LAST as isize,
             Anonymous: windows_sys::Win32::UI::Controls::TVINSERTSTRUCTW_0 { itemex: item },
         };
         let hitem = SendMessageW(state.hwnd_tree, TVM_INSERTITEMW, 0, &mut ins as *mut _ as LPARAM) as isize;
         if first_child == 0 {
             first_child = hitem;
         }
     }

     if first_child != 0 {
         SendMessageW(state.hwnd_tree, TVM_SELECTITEM, TVGN_CARET as usize, first_child as LPARAM);
     }
 }

 unsafe fn open_item_by_index(hwnd: HWND, idx: usize) {
     save_current_if_dirty(hwnd);
 
     let state = state(hwnd);
     if idx >= state.item_paths.len() {
         return;
     }
     let path = state.item_paths[idx].clone();
     match crate::storage::read_text(&path) {
         Ok(content) => {
             let w = wide(&content);
             SetWindowTextW(state.hwnd_edit, w.as_ptr());
             state.current_doc_path = Some(path.clone());
             state.current_doc_dirty = false;
             set_status_text(hwnd, &format!("编辑: {}", path.file_name().and_then(|s| s.to_str()).unwrap_or("")));
         }
         Err(e) => show_error(hwnd, &e.to_string()),
     }
 }

 unsafe fn add_tabs(hwnd_tabs: HWND) {
     let tabs = ["章节", "角色", "世界观", "时间线"];
     for (i, t) in tabs.iter().enumerate() {
         let mut text = wide(t);
         let item = TCITEMW {
             mask: windows_sys::Win32::UI::Controls::TCIF_TEXT as u32,
             dwState: 0,
             dwStateMask: 0,
             pszText: text.as_mut_ptr(),
             cchTextMax: 0,
             iImage: 0,
             lParam: 0,
         };
         SendMessageW(hwnd_tabs, TCM_INSERTITEMW, i, &item as *const _ as LPARAM);
     }
 }

 unsafe fn save_current_if_dirty(hwnd: HWND) {
     let state = state(hwnd);
     if !state.current_doc_dirty {
         return;
     }
     let Some(path) = &state.current_doc_path else { return };
 
     let len = GetWindowTextLengthW(state.hwnd_edit);
     let mut buf = vec![0u16; (len as usize) + 1];
     let read = GetWindowTextW(state.hwnd_edit, buf.as_mut_ptr(), buf.len() as i32);
     if read <= 0 {
         return;
     }
     let s = String::from_utf16_lossy(&buf[..read as usize]);
     if let Err(e) = crate::storage::write_text_atomic(path, &s) {
         show_error(hwnd, &e.to_string());
         return;
     }
    if let Some(project) = &state.project {
        let now = now_unix();
        if now >= state.last_backup_unix.saturating_add(300) {
            let _ = crate::storage::backup_text(&project.root, path, &s);
            state.last_backup_unix = now;
        }
    }
     state.current_doc_dirty = false;
     set_status_text(hwnd, "已自动保存");
 }

unsafe fn get_text(hwnd_ctrl: HWND) -> String {
    if hwnd_ctrl == 0 {
        return String::new();
    }
    let len = GetWindowTextLengthW(hwnd_ctrl);
    if len <= 0 {
        return String::new();
    }
    let mut buf = vec![0u16; (len as usize) + 1];
    let read = GetWindowTextW(hwnd_ctrl, buf.as_mut_ptr(), buf.len() as i32);
    if read <= 0 {
        return String::new();
    }
    String::from_utf16_lossy(&buf[..read as usize])
}

unsafe fn highlight_current_line(hwnd: HWND) {
    let state = state(hwnd);
    let edit = state.hwnd_edit;
    if edit == 0 {
        return;
    }
    let len = GetWindowTextLengthW(edit);
    if len > 200_000 {
        return;
    }

    let mut sel_start: u32 = 0;
    let mut sel_end: u32 = 0;
    SendMessageW(
        edit,
        EM_GETSEL_MSG,
        &mut sel_start as *mut _ as usize,
        &mut sel_end as *mut _ as LPARAM,
    );

    let line = SendMessageW(edit, EM_LINEFROMCHAR_MSG, sel_start as usize, 0) as i32;
    if state.last_highlight_line != -1 && state.last_highlight_line != line {
        apply_line_bold(edit, state.last_highlight_line, false);
    }

    let heading = line_starts_with(edit, line, "#");
    apply_line_bold(edit, line, heading);
    state.last_highlight_line = line;

    SendMessageW(edit, EM_SETSEL_MSG, sel_start as usize, sel_end as LPARAM);
}

unsafe fn line_starts_with(edit: HWND, line: i32, prefix: &str) -> bool {
    let start = SendMessageW(edit, EM_LINEINDEX_MSG, line as usize, 0) as i32;
    if start < 0 {
        return false;
    }
    let line_len = SendMessageW(edit, EM_LINELENGTH_MSG, start as usize, 0) as usize;
    if line_len == 0 {
        return false;
    }
    let max = line_len.min(256);
    let mut buf = vec![0u16; max + 1];
    buf[0] = max as u16;
    let copied = SendMessageW(edit, EM_GETLINE_MSG, line as usize, buf.as_mut_ptr() as LPARAM) as usize;
    let s = String::from_utf16_lossy(&buf[..copied]);
    s.trim_start().starts_with(prefix)
}

unsafe fn apply_line_bold(edit: HWND, line: i32, bold: bool) {
    let start = SendMessageW(edit, EM_LINEINDEX_MSG, line as usize, 0) as i32;
    if start < 0 {
        return;
    }
    let len = SendMessageW(edit, EM_LINELENGTH_MSG, start as usize, 0) as i32;
    if len <= 0 {
        return;
    }
    SendMessageW(edit, EM_SETSEL_MSG, start as usize, (start + len) as LPARAM);

    let mut cf: CHARFORMAT2W = std::mem::zeroed();
    cf.Base.cbSize = std::mem::size_of::<CHARFORMAT2W>() as u32;
    cf.Base.dwMask = CFM_BOLD;
    cf.Base.dwEffects = if bold {
        CFE_BOLD
    } else {
        windows::Win32::UI::Controls::RichEdit::CFE_EFFECTS(0)
    };
    SendMessageW(edit, EM_SETCHARFORMAT, SCF_SELECTION as usize, &cf as *const _ as LPARAM);
}
 
 unsafe fn show_error(owner: HWND, message: &str) {
     let title = wide("错误");
     let msg = wide(message);
     windows_sys::Win32::UI::WindowsAndMessaging::MessageBoxW(
         owner,
         msg.as_ptr(),
         title.as_ptr(),
         windows_sys::Win32::UI::WindowsAndMessaging::MB_ICONERROR,
     );
 }
 
 unsafe fn set_status_text(hwnd: HWND, text: &str) {
    set_status_part(hwnd, 0, text);
}

unsafe fn set_status_part(hwnd: HWND, part: usize, text: &str) {
    let state = state(hwnd);
    if state.hwnd_status == 0 {
        return;
    }
    let t = wide(text);
    SendMessageW(state.hwnd_status, SB_SETTEXTW, part, t.as_ptr() as LPARAM);
 }
 
unsafe fn apply_theme(hwnd: HWND, dark: bool) {
    let value: u32 = if dark { 1 } else { 0 };
    let _ = DwmSetWindowAttribute(
        hwnd,
        DWMWA_USE_IMMERSIVE_DARK_MODE,
        &value as *const _ as *const c_void,
        std::mem::size_of_val(&value) as u32,
    );

    let state = state(hwnd);
    let (bg, fg, line) = if dark {
        (0x00202020u32, 0x00E0E0E0u32, 0x00303030u32)
    } else {
        (0x00FFFFFFu32, 0x00000000u32, 0x00C0C0C0u32)
    };
    let theme = if dark { wide("DarkMode_Explorer") } else { wide("Explorer") };

    if state.hwnd_tabs != 0 {
        let _ = SetWindowTheme(state.hwnd_tabs, theme.as_ptr(), null());
    }
    if state.hwnd_search != 0 {
        let _ = SetWindowTheme(state.hwnd_search, theme.as_ptr(), null());
    }
    if state.hwnd_tree != 0 {
        let _ = SetWindowTheme(state.hwnd_tree, theme.as_ptr(), null());
        SendMessageW(state.hwnd_tree, TVM_SETBKCOLOR, 0, bg as LPARAM);
        SendMessageW(state.hwnd_tree, TVM_SETTEXTCOLOR, 0, fg as LPARAM);
        SendMessageW(state.hwnd_tree, TVM_SETLINECOLOR, 0, line as LPARAM);
    }
    if state.hwnd_status != 0 {
        let _ = SetWindowTheme(state.hwnd_status, theme.as_ptr(), null());
    }
    if state.hwnd_splitter != 0 {
        let _ = SetWindowTheme(state.hwnd_splitter, theme.as_ptr(), null());
    }
    if state.hwnd_edit != 0 {
        SendMessageW(state.hwnd_edit, EM_SETBKGNDCOLOR_MSG, 0, bg as LPARAM);
        let mut cf: CHARFORMAT2W = std::mem::zeroed();
        cf.Base.cbSize = std::mem::size_of::<CHARFORMAT2W>() as u32;
        cf.Base.dwMask = CFM_COLOR;
        if dark {
            cf.Base.crTextColor = windows::Win32::Foundation::COLORREF(fg);
            cf.Base.dwEffects = windows::Win32::UI::Controls::RichEdit::CFE_EFFECTS(0);
        } else {
            cf.Base.dwEffects = CFE_AUTOCOLOR;
        }
        SendMessageW(state.hwnd_edit, EM_SETCHARFORMAT, SCF_DEFAULT as usize, &cf as *const _ as LPARAM);
    }
    InvalidateRect(hwnd, null(), 1);
}

unsafe fn check_updates(hwnd: HWND) -> Result<(), String> {
    let state = state(hwnd);
    let Some(project) = &state.project else {
        return Err("未打开项目".to_string());
    };
    let Some(url) = project.meta.update_url.as_deref() else {
        return Err("请在project.md的front matter中配置update_url".to_string());
    };
    let current = env!("CARGO_PKG_VERSION");
    match crate::update::check_update(current, url)? {
        Some(info) => {
            let mut msg = format!("发现新版本: {}\n当前版本: {}", info.version, current);
            if let Some(u) = info.url.as_deref() {
                msg.push_str("\n下载: ");
                msg.push_str(u);
            }
            if let Some(notes) = info.notes {
                msg.push_str("\n\n");
                msg.push_str(&notes);
            }
            show_error(hwnd, &msg);
            Ok(())
        }
        None => {
            set_status_text(hwnd, "已是最新版本");
            Ok(())
        }
    }
}

 unsafe fn pick_folder(owner: HWND) -> Option<PathBuf> {
     use windows::Win32::Foundation::HWND as WndHwnd;
     use windows::Win32::System::Com::{CoCreateInstance, CoTaskMemFree, CLSCTX_INPROC_SERVER};
     use windows::Win32::UI::Shell::{
         FileOpenDialog, IFileOpenDialog, FOS_FORCEFILESYSTEM, FOS_PICKFOLDERS, SIGDN_FILESYSPATH,
     };
 
     let dialog: IFileOpenDialog = CoCreateInstance(&FileOpenDialog, None, CLSCTX_INPROC_SERVER).ok()?;
     let options = dialog.GetOptions().ok()?;
     dialog
         .SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM)
         .ok()?;
     let _ = dialog.SetTitle(windows::core::w!("选择或创建项目文件夹"));
     dialog.Show(WndHwnd(owner)).ok()?;
 
     let item = dialog.GetResult().ok()?;
     let p = item.GetDisplayName(SIGDN_FILESYSPATH).ok()?;
     let s = p.to_string().ok()?;
     CoTaskMemFree(Some(p.0 as _));
     Some(PathBuf::from(s))
 }

unsafe fn export_project(hwnd: HWND) -> Result<(), String> {
    let state = state(hwnd);
    let Some(project) = &state.project else { return Ok(()) };
    let Some(dst_root) = pick_folder(hwnd) else { return Ok(()) };
    let ts = now_unix();
    let name = sanitize_filename(&project.meta.name);
    let base = if name.is_empty() { "project".to_string() } else { name };
    let dst = dst_root.join(format!("{}-export-{}", base, ts));
    copy_project(&project.root, &dst).map_err(|e| e.to_string())?;
    set_status_text(hwnd, &format!("已导出: {}", dst.display()));
    Ok(())
}

unsafe fn import_project(hwnd: HWND) -> Result<(), String> {
    let Some(src_root) = pick_folder(hwnd) else { return Ok(()) };
    let Some(dst_parent) = pick_folder(hwnd) else { return Ok(()) };
    let ts = now_unix();
    let src_name = src_root.file_name().and_then(|s| s.to_str()).unwrap_or("import");
    let dst = dst_parent.join(format!("{}-import-{}", sanitize_filename(src_name), ts));
    copy_project(&src_root, &dst).map_err(|e| e.to_string())?;

    match ProjectStore::open_or_init(dst) {
        Ok(project) => {
            let state = state(hwnd);
            state.project = Some(project);
            state.dark_mode = state
                .project
                .as_ref()
                .and_then(|p| p.meta.theme.as_deref())
                .map(|t| t.eq_ignore_ascii_case("dark"))
                .unwrap_or(false);
            apply_theme(hwnd, state.dark_mode);
            state.left_pane_ratio = state
                .project
                .as_ref()
                .and_then(|p| p.meta.left_pane_ratio)
                .unwrap_or(0.28);
            state.current_section = Section::Chapters;
            state.item_paths.clear();
            state.current_doc_path = None;
            state.current_doc_dirty = false;
            state.filter_text.clear();
            state.last_highlight_line = -1;
            state.undo_stack.clear();
            state.redo_stack.clear();
            state.dragging = false;
            state.last_backup_unix = 0;
            SetWindowTextW(state.hwnd_search, wide("").as_ptr());
            SendMessageW(state.hwnd_tree, TVM_DELETEITEM, 0, 0);
            SendMessageW(state.hwnd_tabs, windows_sys::Win32::UI::Controls::TCM_SETCURSEL, 0, 0);
            reload_items(hwnd);
            let _ = SetTimer(hwnd, TIMER_AUTOSAVE, 30_000, None);
            if let Some(project) = &state.project {
                set_status_text(hwnd, &format!("已导入并打开: {}", project.root.display()));
                set_status_part(hwnd, 1, "");
                set_status_part(hwnd, 2, "已保存");
            }
            Ok(())
        }
        Err(e) => Err(e),
    }
}

fn copy_project(src: &PathBuf, dst: &PathBuf) -> std::io::Result<()> {
    std::fs::create_dir_all(dst)?;
    let files = ["project.md"];
    for f in files {
        let from = src.join(f);
        if from.exists() {
            let to = dst.join(f);
            std::fs::copy(from, to)?;
        }
    }
    for dir in ["chapters", "characters", "world", "timeline"] {
        let from_dir = src.join(dir);
        if !from_dir.exists() {
            continue;
        }
        let to_dir = dst.join(dir);
        copy_dir_recursive(&from_dir, &to_dir)?;
    }
    Ok(())
}

fn copy_dir_recursive(src: &PathBuf, dst: &PathBuf) -> std::io::Result<()> {
    std::fs::create_dir_all(dst)?;
    for entry in std::fs::read_dir(src)? {
        let entry = entry?;
        let path = entry.path();
        let name = entry.file_name();
        let to = dst.join(name);
        if path.is_dir() {
            copy_dir_recursive(&path, &to)?;
        } else if path.is_file() {
            std::fs::copy(path, to)?;
        }
    }
    Ok(())
}

fn now_unix() -> u64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0)
}

fn apply_rename_pairs(pairs: &[(PathBuf, PathBuf)]) -> std::io::Result<()> {
    if pairs.is_empty() {
        return Ok(());
    }
    let pid = std::process::id();
    let mut temp_pairs: Vec<(PathBuf, PathBuf)> = Vec::with_capacity(pairs.len());
    for (i, (from, _to)) in pairs.iter().enumerate() {
        let file = from.file_name().and_then(|s| s.to_str()).unwrap_or("file");
        let tmp_name = format!("{}.reorder.tmp.{}.{}", file, pid, i);
        let tmp = from.with_file_name(tmp_name);
        std::fs::rename(from, &tmp)?;
        temp_pairs.push((tmp, pairs[i].1.clone()));
    }
    for (tmp, to) in temp_pairs {
        std::fs::rename(tmp, to)?;
    }
    Ok(())
}

unsafe fn begin_rename_selected(hwnd: HWND) {
    let state = state(hwnd);
    if state.hwnd_tree == 0 {
        return;
    }
    let hitem = SendMessageW(state.hwnd_tree, TVM_GETNEXTITEM, TVGN_CARET as usize, 0) as isize;
    if hitem == 0 {
        return;
    }
    SendMessageW(state.hwnd_tree, TVM_EDITLABELW, 0, hitem as LPARAM);
}

unsafe fn commit_rename(hwnd: HWND, idx: usize, new_name: &str) -> Result<bool, String> {
    let state = state(hwnd);
    if idx >= state.item_paths.len() {
        return Ok(false);
    }
    let Some(project) = &state.project else { return Ok(false) };
    let from = state.item_paths[idx].clone();

    let name = new_name.trim();
    if name.is_empty() {
        return Ok(false);
    }

    let to = match state.current_section {
        Section::Chapters => {
            let stem = from.file_stem().and_then(|s| s.to_str()).unwrap_or("");
            let prefix = if stem.len() >= 5 && stem.as_bytes()[4] == b'-' && stem[..4].chars().all(|c| c.is_ascii_digit()) {
                &stem[..5]
            } else {
                ""
            };
            let mut base = sanitize_filename(name);
            if base.ends_with(".md") {
                base.truncate(base.len() - 3);
            }
            let file = format!("{}{}.md", prefix, base);
            project.root.join("chapters").join(unique_file_name(&project.root.join("chapters"), &file))
        }
        Section::Characters => project
            .root
            .join("characters")
            .join(unique_file_name(&project.root.join("characters"), &format!("{}.md", sanitize_filename(name)))),
        Section::World => project
            .root
            .join("world")
            .join(unique_file_name(&project.root.join("world"), &format!("{}.md", sanitize_filename(name)))),
        Section::Timeline => project
            .root
            .join("timeline")
            .join(unique_file_name(&project.root.join("timeline"), &format!("{}.md", sanitize_filename(name)))),
    };

    if from == to {
        return Ok(false);
    }

    let cmd = Command::RenameFile { from: from.clone(), to: to.clone() };
    cmd.apply()?;
    state.undo_stack.push(cmd);
    state.redo_stack.clear();
    if state.current_doc_path.as_ref() == Some(&from) {
        state.current_doc_path = Some(to.clone());
    }
    SendMessageW(state.hwnd_tree, TVM_DELETEITEM, 0, 0);
    reload_items(hwnd);
    Ok(true)
}

unsafe fn create_new_item(hwnd: HWND) -> Result<(), String> {
    let state = state(hwnd);
    let Some(project) = &state.project else { return Ok(()) };

    state.filter_text.clear();
    SetWindowTextW(state.hwnd_search, wide("").as_ptr());

    let (dir, default_title) = match state.current_section {
        Section::Chapters => (project.root.join("chapters"), "新建章节"),
        Section::Characters => (project.root.join("characters"), "新建角色"),
        Section::World => (project.root.join("world"), "新建设定"),
        Section::Timeline => (project.root.join("timeline"), "新建时间线条目"),
    };

    let path = if state.current_section == Section::Chapters {
        let existing = crate::storage::list_markdown_files(&dir).map_err(|e| e.to_string())?;
        let next_num = existing
            .iter()
            .filter_map(|p| p.file_stem().and_then(|s| s.to_str()))
            .filter_map(|s| s.get(0..4).and_then(|n| n.parse::<u32>().ok()))
            .max()
            .unwrap_or(0)
            + 1;
        let file = format!("{:04}-{}.md", next_num, default_title);
        dir.join(unique_file_name(&dir, &file))
    } else {
        dir.join(unique_file_name(&dir, &format!("{}.md", default_title)))
    };

    let contents = format!("---\nformat_version: 1\n---\n\n# {}\n\n", default_title);
    let cmd = Command::CreateFile { path: path.clone(), contents };
    cmd.apply()?;
    state.undo_stack.push(cmd);
    state.redo_stack.clear();

    SendMessageW(state.hwnd_tree, TVM_DELETEITEM, 0, 0);
    reload_items(hwnd);
    if let Some(i) = state.item_paths.iter().position(|p| p == &path) {
        open_item_by_index(hwnd, i);
    }
    Ok(())
}

unsafe fn delete_selected_item(hwnd: HWND) -> Result<(), String> {
    save_current_if_dirty(hwnd);
    let state = state(hwnd);
    let Some(path) = state.current_doc_path.clone() else { return Ok(()) };
    let contents = crate::storage::read_text(&path).unwrap_or_default();
    let cmd = Command::DeleteFile { path: path.clone(), contents };
    cmd.apply()?;
    state.undo_stack.push(cmd);
    state.redo_stack.clear();
    state.current_doc_path = None;
    state.current_doc_dirty = false;
    SendMessageW(state.hwnd_tree, TVM_DELETEITEM, 0, 0);
    reload_items(hwnd);
    Ok(())
}

unsafe fn do_undo(hwnd: HWND) -> Result<(), String> {
    save_current_if_dirty(hwnd);
    let state = state(hwnd);
    let Some(cmd) = state.undo_stack.pop() else { return Ok(()) };
    cmd.undo()?;
    state.redo_stack.push(cmd);
    SendMessageW(state.hwnd_tree, TVM_DELETEITEM, 0, 0);
    reload_items(hwnd);
    Ok(())
}

unsafe fn do_redo(hwnd: HWND) -> Result<(), String> {
    save_current_if_dirty(hwnd);
    let state = state(hwnd);
    let Some(cmd) = state.redo_stack.pop() else { return Ok(()) };
    cmd.apply()?;
    state.undo_stack.push(cmd);
    SendMessageW(state.hwnd_tree, TVM_DELETEITEM, 0, 0);
    reload_items(hwnd);
    Ok(())
}

unsafe fn drag_update_hover(hwnd: HWND) {
    let state = state(hwnd);
    if state.hwnd_tree == 0 {
        return;
    }
    let Some(hitem) = tree_item_at_cursor(state.hwnd_tree) else { return };
    SendMessageW(state.hwnd_tree, TVM_SELECTITEM, TVGN_CARET as usize, hitem as LPARAM);
}

unsafe fn drag_commit_drop(hwnd: HWND, src_idx: usize) -> Result<(), String> {
    let state = state(hwnd);
    if state.current_section != Section::Chapters {
        return Ok(());
    }
    if !state.filter_text.trim().is_empty() {
        return Ok(());
    }
    let Some(project) = &state.project else { return Ok(()) };
    if src_idx >= state.item_paths.len() {
        return Ok(());
    }
    let Some(dst_idx) = tree_index_at_cursor(state.hwnd_tree) else { return Ok(()) };
    if dst_idx == src_idx {
        return Ok(());
    }

    let chapters_dir = project.root.join("chapters");
    let mut all = crate::storage::list_markdown_files(&chapters_dir).map_err(|e| e.to_string())?;
    let src_path = state.item_paths[src_idx].clone();
    let dst_path = state.item_paths[dst_idx].clone();
    let src_pos = all.iter().position(|p| p == &src_path).ok_or_else(|| "source not found".to_string())?;
    let dst_pos = all.iter().position(|p| p == &dst_path).ok_or_else(|| "target not found".to_string())?;

    let moved = all.remove(src_pos);
    let insert_at = if src_pos < dst_pos { dst_pos - 1 } else { dst_pos };
    all.insert(insert_at, moved);

    let mut pairs = Vec::with_capacity(all.len());
    for (i, old) in all.iter().enumerate() {
        let stem = old.file_stem().and_then(|s| s.to_str()).unwrap_or("");
        let base = if stem.len() >= 5 && stem.as_bytes()[4] == b'-' && stem[..4].chars().all(|c| c.is_ascii_digit()) {
            &stem[5..]
        } else {
            stem
        };
        let base = if base.trim().is_empty() { "章节" } else { base };
        let file = format!("{:04}-{}.md", i + 1, sanitize_filename(base));
        let new_path = chapters_dir.join(file);
        pairs.push((old.clone(), new_path));
    }

    let cmd = Command::ReorderChapters { pairs };
    cmd.apply()?;
    state.undo_stack.push(cmd);
    state.redo_stack.clear();
    SendMessageW(state.hwnd_tree, TVM_DELETEITEM, 0, 0);
    reload_items(hwnd);
    Ok(())
}

unsafe fn tree_item_at_cursor(hwnd_tree: HWND) -> Option<isize> {
    let mut pt = windows_sys::Win32::Foundation::POINT { x: 0, y: 0 };
    if GetCursorPos(&mut pt) == 0 {
        return None;
    }
    ScreenToClient(hwnd_tree, &mut pt);
    let mut hti: TVHITTESTINFO = std::mem::zeroed();
    hti.pt = pt;
    SendMessageW(hwnd_tree, TVM_HITTEST, 0, &mut hti as *mut _ as LPARAM);
    if hti.hItem == 0 {
        None
    } else {
        Some(hti.hItem as isize)
    }
}

unsafe fn tree_index_at_cursor(hwnd_tree: HWND) -> Option<usize> {
    let hitem = tree_item_at_cursor(hwnd_tree)?;
    let mut item: TVITEMEXW = std::mem::zeroed();
    item.mask = TVIF_PARAM as u32;
    item.hItem = hitem;
    let ok = SendMessageW(hwnd_tree, TVM_GETITEMW, 0, &mut item as *mut _ as LPARAM);
    if ok == 0 {
        return None;
    }
    let idx = item.lParam as isize;
    if idx < 0 {
        None
    } else {
        Some(idx as usize)
    }
}

fn unique_file_name(dir: &PathBuf, file_name: &str) -> String {
    if !dir.join(file_name).exists() {
        return file_name.to_string();
    }
    let (stem, ext) = match file_name.rsplit_once('.') {
        Some((s, e)) => (s.to_string(), format!(".{}", e)),
        None => (file_name.to_string(), String::new()),
    };
    for i in 1..10_000 {
        let cand = format!("{} ({}){}", stem, i, ext);
        if !dir.join(&cand).exists() {
            return cand;
        }
    }
    file_name.to_string()
}

fn sanitize_filename(name: &str) -> String {
    let mut out = String::new();
    for c in name.chars() {
        if matches!(c, '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*') {
            out.push('_');
        } else {
            out.push(c);
        }
    }
    out.trim().trim_matches('.').to_string()
}

unsafe fn wide_ptr_to_string(ptr: *mut u16) -> String {
    if ptr.is_null() {
        return String::new();
    }
    let mut len = 0usize;
    loop {
        let c = *ptr.add(len);
        if c == 0 {
            break;
        }
        len += 1;
        if len > 32_768 {
            break;
        }
    }
    let slice = std::slice::from_raw_parts(ptr, len);
    String::from_utf16_lossy(slice)
}
