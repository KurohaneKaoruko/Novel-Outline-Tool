use serde::Deserialize;
use std::ptr::{null, null_mut};

use windows_sys::Win32::Foundation::GetLastError;
use windows_sys::Win32::Networking::WinHttp::{
    WinHttpCloseHandle, WinHttpConnect, WinHttpOpen, WinHttpOpenRequest, WinHttpReadData,
    WinHttpReceiveResponse, WinHttpSendRequest, WINHTTP_ACCESS_TYPE_NO_PROXY,
    WINHTTP_FLAG_SECURE,
};

#[derive(Debug, Clone, Deserialize)]
pub struct UpdateInfo {
    pub version: String,
    #[serde(default)]
    pub url: Option<String>,
    #[serde(default)]
    pub notes: Option<String>,
}

pub fn check_update(current_version: &str, update_json_url: &str) -> Result<Option<UpdateInfo>, String> {
    let body = fetch_text(update_json_url)?;
    let info: UpdateInfo = serde_json::from_str(&body).map_err(|e| e.to_string())?;
    if version_gt(&info.version, current_version) {
        Ok(Some(info))
    } else {
        Ok(None)
    }
}

fn fetch_text(url: &str) -> Result<String, String> {
    let (host, port, path, secure) = parse_url(url).ok_or_else(|| "unsupported update url".to_string())?;
    if !secure {
        return Err("only https is supported".to_string());
    }

    unsafe {
        let agent = wide("NovelOutlineTool");
        let h_session = WinHttpOpen(
            agent.as_ptr(),
            WINHTTP_ACCESS_TYPE_NO_PROXY,
            null(),
            null(),
            0,
        );
        if h_session.is_null() {
            return Err(format!("WinHttpOpen failed: {}", GetLastError()));
        }

        let host_w = wide(&host);
        let h_connect = WinHttpConnect(h_session, host_w.as_ptr(), port, 0);
        if h_connect.is_null() {
            let err = GetLastError();
            WinHttpCloseHandle(h_session);
            return Err(format!("WinHttpConnect failed: {}", err));
        }

        let verb = wide("GET");
        let path_w = wide(&path);
        let h_request = WinHttpOpenRequest(
            h_connect,
            verb.as_ptr(),
            path_w.as_ptr(),
            null(),
            null(),
            null(),
            WINHTTP_FLAG_SECURE,
        );
        if h_request.is_null() {
            let err = GetLastError();
            WinHttpCloseHandle(h_connect);
            WinHttpCloseHandle(h_session);
            return Err(format!("WinHttpOpenRequest failed: {}", err));
        }

        let ok = WinHttpSendRequest(h_request, null(), 0, null_mut(), 0, 0, 0);
        if ok == 0 {
            let err = GetLastError();
            WinHttpCloseHandle(h_request);
            WinHttpCloseHandle(h_connect);
            WinHttpCloseHandle(h_session);
            return Err(format!("WinHttpSendRequest failed: {}", err));
        }

        let ok = WinHttpReceiveResponse(h_request, null_mut());
        if ok == 0 {
            let err = GetLastError();
            WinHttpCloseHandle(h_request);
            WinHttpCloseHandle(h_connect);
            WinHttpCloseHandle(h_session);
            return Err(format!("WinHttpReceiveResponse failed: {}", err));
        }

        let mut out: Vec<u8> = Vec::new();
        let mut buf = vec![0u8; 16 * 1024];
        loop {
            let mut read: u32 = 0;
            let ok = WinHttpReadData(h_request, buf.as_mut_ptr() as *mut _, buf.len() as u32, &mut read);
            if ok == 0 {
                let err = GetLastError();
                WinHttpCloseHandle(h_request);
                WinHttpCloseHandle(h_connect);
                WinHttpCloseHandle(h_session);
                return Err(format!("WinHttpReadData failed: {}", err));
            }
            if read == 0 {
                break;
            }
            out.extend_from_slice(&buf[..read as usize]);
        }

        WinHttpCloseHandle(h_request);
        WinHttpCloseHandle(h_connect);
        WinHttpCloseHandle(h_session);

        String::from_utf8(out).map_err(|e| e.to_string())
    }
}

fn parse_url(url: &str) -> Option<(String, u16, String, bool)> {
    let url = url.trim();
    let secure = url.starts_with("https://");
    let rest = if secure {
        &url["https://".len()..]
    } else if url.starts_with("http://") {
        &url["http://".len()..]
    } else {
        return None;
    };

    let (host_port, path) = match rest.split_once('/') {
        Some((h, p)) => (h, format!("/{}", p)),
        None => (rest, "/".to_string()),
    };

    let (host, port) = match host_port.split_once(':') {
        Some((h, p)) => (h.to_string(), p.parse::<u16>().ok()?),
        None => (host_port.to_string(), if secure { 443 } else { 80 }),
    };

    Some((host, port, path, secure))
}

fn version_gt(a: &str, b: &str) -> bool {
    let pa = parse_version(a);
    let pb = parse_version(b);
    pa > pb
}

fn parse_version(v: &str) -> (u32, u32, u32) {
    let mut it = v.split(|c| c == '.' || c == '-' || c == '+');
    let major = it.next().and_then(|s| s.parse().ok()).unwrap_or(0);
    let minor = it.next().and_then(|s| s.parse().ok()).unwrap_or(0);
    let patch = it.next().and_then(|s| s.parse().ok()).unwrap_or(0);
    (major, minor, patch)
}

fn wide(s: &str) -> Vec<u16> {
    let mut v: Vec<u16> = s.encode_utf16().collect();
    v.push(0);
    v
}
