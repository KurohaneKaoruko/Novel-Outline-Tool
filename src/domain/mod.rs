 use serde::{Deserialize, Serialize};
 use std::path::PathBuf;
 use std::time::{SystemTime, UNIX_EPOCH};
 
 #[derive(Debug, Clone, Serialize, Deserialize)]
 pub struct ProjectMeta {
     pub name: String,
     pub created_unix: u64,
     pub format_version: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub update_url: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left_pane_ratio: Option<f32>,
 }
 
 impl ProjectMeta {
     pub fn new(name: String) -> Self {
         Self {
             name,
             created_unix: now_unix(),
             format_version: 1,
            update_url: None,
            theme: None,
            left_pane_ratio: None,
         }
     }
 }
 
 impl Default for ProjectMeta {
     fn default() -> Self {
         Self {
             name: "Untitled".to_string(),
             created_unix: now_unix(),
             format_version: 1,
            update_url: None,
            theme: None,
            left_pane_ratio: None,
         }
     }
 }
 
 #[derive(Debug, Clone)]
 pub struct Project {
     pub root: PathBuf,
     pub meta: ProjectMeta,
 }
 
 fn now_unix() -> u64 {
     SystemTime::now()
         .duration_since(UNIX_EPOCH)
         .map(|d| d.as_secs())
         .unwrap_or(0)
 }
