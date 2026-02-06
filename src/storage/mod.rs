 use crate::domain::{Project, ProjectMeta};
 use serde::de::DeserializeOwned;
 use std::fs;
 use std::io;
 use std::os::windows::ffi::OsStrExt;
 use std::path::{Path, PathBuf};
 
 use windows_sys::Win32::Storage::FileSystem::{
     MoveFileExW, MOVEFILE_REPLACE_EXISTING, MOVEFILE_WRITE_THROUGH,
 };
 
 pub struct ProjectStore {
     root: PathBuf,
 }
 
pub fn list_markdown_files(dir: &Path) -> io::Result<Vec<PathBuf>> {
    let mut out = Vec::new();
    for entry in fs::read_dir(dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.is_file() && path.extension().and_then(|e| e.to_str()).map(|e| e.eq_ignore_ascii_case("md")).unwrap_or(false) {
            out.push(path);
        }
    }
    out.sort_by(|a, b| a.file_name().cmp(&b.file_name()));
    Ok(out)
}

pub fn read_text(path: &Path) -> io::Result<String> {
    fs::read_to_string(path)
}

pub fn write_text_atomic(path: &Path, text: &str) -> io::Result<()> {
    atomic_write(path, text.as_bytes())
}

pub fn cleanup_temp_files(project_root: &Path) -> io::Result<()> {
    let dirs = [
        project_root.to_path_buf(),
        project_root.join("chapters"),
        project_root.join("characters"),
        project_root.join("world"),
        project_root.join("timeline"),
    ];
    for dir in dirs {
        if !dir.exists() {
            continue;
        }
        for entry in fs::read_dir(dir)? {
            let entry = entry?;
            let path = entry.path();
            if !path.is_file() {
                continue;
            }
            let name = path.file_name().and_then(|s| s.to_str()).unwrap_or("");
            if name.contains(".tmp.") {
                let _ = fs::remove_file(path);
            }
        }
    }
    Ok(())
}

pub fn backup_text(project_root: &Path, source_path: &Path, text: &str) -> io::Result<()> {
    let backup_root = project_root.join(".backup");
    ensure_dir(&backup_root)?;

    let ts = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs())
        .unwrap_or(0);
    let snapshot_dir = backup_root.join(format!("backup-{}", ts));
    ensure_dir(&snapshot_dir)?;

    let rel = source_path.strip_prefix(project_root).unwrap_or(source_path);
    let dst = snapshot_dir.join(rel);
    if let Some(parent) = dst.parent() {
        ensure_dir(parent)?;
    }
    atomic_write(&dst, text.as_bytes())?;

    cleanup_old_backups(&backup_root, 30)?;
    Ok(())
}

 impl ProjectStore {
     pub fn open_or_init(root: PathBuf) -> Result<Project, String> {
         let store = Self { root };
         store.ensure_layout().map_err(|e| e.to_string())?;
        let _ = cleanup_temp_files(&store.root);
 
         let project_md = store.root.join("project.md");
         if !project_md.exists() {
             let name = store
                 .root
                 .file_name()
                 .and_then(|s| s.to_str())
                 .unwrap_or("Untitled")
                 .to_string();
             let meta = ProjectMeta::new(name);
             store
                 .write_project_md(&meta, default_project_body())
                 .map_err(|e| e.to_string())?;
             return Ok(Project {
                 root: store.root,
                 meta,
             });
         }
 
         let (meta, _body) = store.read_project_md().map_err(|e| e.to_string())?;
         Ok(Project {
             root: store.root,
             meta,
         })
     }
 
     pub fn save_project_meta(project: &Project) -> Result<(), String> {
         let store = Self {
             root: project.root.clone(),
         };
         let (_old_meta, body) = store.read_project_md().unwrap_or((ProjectMeta::default(), String::new()));
         store
             .write_project_md(&project.meta, if body.trim().is_empty() { default_project_body() } else { body })
             .map_err(|e| e.to_string())
     }
 
     fn ensure_layout(&self) -> io::Result<()> {
         ensure_dir(&self.root)?;
        let chapters_dir = self.root.join("chapters");
        ensure_dir(&chapters_dir)?;
        let characters_dir = self.root.join("characters");
        ensure_dir(&characters_dir)?;
        let world_dir = self.root.join("world");
        ensure_dir(&world_dir)?;
        let timeline_dir = self.root.join("timeline");
        ensure_dir(&timeline_dir)?;
         ensure_dir(&self.root.join("exports"))?;
         ensure_dir(&self.root.join(".backup"))?;
 
        if list_markdown_files(&chapters_dir)?.is_empty() {
            let first = chapters_dir.join("0001-第一章.md");
            atomic_write(&first, default_chapter_md().as_bytes())?;
        }

        if list_markdown_files(&characters_dir)?.is_empty() {
            let first = characters_dir.join("角色-示例.md");
            atomic_write(&first, default_character_md().as_bytes())?;
        }

        if list_markdown_files(&world_dir)?.is_empty() {
            let first = world_dir.join("世界观-示例.md");
            atomic_write(&first, default_world_md().as_bytes())?;
        }

        let timeline_md = timeline_dir.join("timeline.md");
         if !timeline_md.exists() {
             atomic_write(&timeline_md, default_timeline_md().as_bytes())?;
         }
 
         Ok(())
     }
 
     fn read_project_md(&self) -> io::Result<(ProjectMeta, String)> {
         let path = self.root.join("project.md");
         let content = fs::read_to_string(path)?;
         let (meta, body) = parse_front_matter::<ProjectMeta>(&content);
         Ok((meta.unwrap_or_default(), body.to_string()))
     }
 
     fn write_project_md(&self, meta: &ProjectMeta, body: String) -> io::Result<()> {
         let yaml = serde_yaml::to_string(meta).unwrap_or_default();
         let mut out = String::new();
         out.push_str("---\n");
         out.push_str(yaml.trim_end());
         out.push('\n');
         out.push_str("---\n\n");
         out.push_str(body.trim_start());
         out.push('\n');
 
         let path = self.root.join("project.md");
         atomic_write(&path, out.as_bytes())
     }
 }

fn cleanup_old_backups(backup_root: &Path, keep: usize) -> io::Result<()> {
    let mut dirs: Vec<PathBuf> = Vec::new();
    for entry in fs::read_dir(backup_root)? {
        let entry = entry?;
        let path = entry.path();
        if path.is_dir() {
            let name = path.file_name().and_then(|s| s.to_str()).unwrap_or("");
            if name.starts_with("backup-") {
                dirs.push(path);
            }
        }
    }
    dirs.sort();
    if dirs.len() <= keep {
        return Ok(());
    }
    let remove = dirs.len() - keep;
    for path in dirs.into_iter().take(remove) {
        let _ = fs::remove_dir_all(path);
    }
    Ok(())
}
 
 fn ensure_dir(path: &Path) -> io::Result<()> {
     fs::create_dir_all(path)
 }
 
 fn parse_front_matter<T: DeserializeOwned>(content: &str) -> (Option<T>, &str) {
     let normalized = content.strip_prefix("\u{feff}").unwrap_or(content);
     let content = normalized;
 
     if !content.starts_with("---\n") && !content.starts_with("---\r\n") {
         return (None, content);
     }
 
     let rest = &content[3..];
     let rest = rest.strip_prefix("\r\n").or_else(|| rest.strip_prefix("\n")).unwrap_or(rest);
 
     if let Some(end) = find_front_matter_end(rest) {
         let (yaml_part, body_part) = rest.split_at(end);
         let yaml = yaml_part.trim();
         let body = body_part
             .trim_start_matches("\r\n---\r\n")
             .trim_start_matches("\n---\n")
             .trim_start_matches("\r\n---\n")
             .trim_start_matches("\n---\r\n");
 
         let meta = serde_yaml::from_str::<T>(yaml).ok();
         (meta, body)
     } else {
         (None, content)
     }
 }
 
 fn find_front_matter_end(s: &str) -> Option<usize> {
     let patterns = ["\n---\n", "\r\n---\r\n", "\n---\r\n", "\r\n---\n"];
     patterns.iter().filter_map(|p| s.find(p)).min()
 }
 
 fn atomic_write(path: &Path, bytes: &[u8]) -> io::Result<()> {
     let tmp = temp_path(path);
     fs::write(&tmp, bytes)?;
     replace_file(&tmp, path)?;
     Ok(())
 }
 
 fn temp_path(path: &Path) -> PathBuf {
     let mut tmp = path.as_os_str().to_owned();
     tmp.push(format!(".tmp.{}", std::process::id()));
     PathBuf::from(tmp)
 }
 
 fn replace_file(src: &Path, dst: &Path) -> io::Result<()> {
     let src_w = to_wide_path(src);
     let dst_w = to_wide_path(dst);
     let ok = unsafe {
         MoveFileExW(
             src_w.as_ptr(),
             dst_w.as_ptr(),
             MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH,
         )
     };
     if ok == 0 {
         let _ = fs::remove_file(dst);
         fs::rename(src, dst)?;
     }
     Ok(())
 }
 
 fn to_wide_path(path: &Path) -> Vec<u16> {
     let mut buf: Vec<u16> = path.as_os_str().encode_wide().collect();
     buf.push(0);
     buf
 }
 
 fn default_project_body() -> String {
     "## 概览\n\n- 这是项目根说明文件。\n- 左侧创建章节/角色/设定/时间线条目，右侧编辑对应 Markdown。\n"
         .to_string()
 }
 
 fn default_timeline_md() -> String {
     "---\nformat_version: 1\n---\n\n# 时间线\n\n| 时间 | 事件 | 关联 |\n| --- | --- | --- |\n"
         .to_string()
 }

fn default_chapter_md() -> String {
    "---\nformat_version: 1\n---\n\n# 第一章\n\n"
        .to_string()
}

fn default_character_md() -> String {
    "---\nformat_version: 1\n---\n\n# 角色：示例\n\n- 角色定位：\n- 外貌特征：\n- 性格：\n- 目标与动机：\n- 关系：\n"
        .to_string()
}

fn default_world_md() -> String {
    "---\nformat_version: 1\n---\n\n# 世界观：示例\n\n## 规则\n\n## 地理\n\n## 势力\n\n"
        .to_string()
}
