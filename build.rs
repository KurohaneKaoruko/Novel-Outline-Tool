fn main() {
    use std::{env, fs, path::PathBuf};

    fn push_u16(buf: &mut Vec<u8>, v: u16) {
        buf.extend_from_slice(&v.to_le_bytes());
    }

    fn push_u32(buf: &mut Vec<u8>, v: u32) {
        buf.extend_from_slice(&v.to_le_bytes());
    }

    fn build_ico_16_rgba(r: u8, g: u8, b: u8, a: u8) -> Vec<u8> {
        let w = 16u32;
        let h = 16u32;
        let and_row_bytes = 4u32;
        let and_bytes = and_row_bytes * h;
        let xor_bytes = w * h * 4;
        let dib_bytes = 40u32 + xor_bytes + and_bytes;
        let offset = 6u32 + 16u32;

        let mut out = Vec::with_capacity((offset + dib_bytes) as usize);

        push_u16(&mut out, 0);
        push_u16(&mut out, 1);
        push_u16(&mut out, 1);

        out.push(16);
        out.push(16);
        out.push(0);
        out.push(0);
        push_u16(&mut out, 1);
        push_u16(&mut out, 32);
        push_u32(&mut out, dib_bytes);
        push_u32(&mut out, offset);

        push_u32(&mut out, 40);
        push_u32(&mut out, w);
        push_u32(&mut out, h * 2);
        push_u16(&mut out, 1);
        push_u16(&mut out, 32);
        push_u32(&mut out, 0);
        push_u32(&mut out, xor_bytes);
        push_u32(&mut out, 0);
        push_u32(&mut out, 0);
        push_u32(&mut out, 0);
        push_u32(&mut out, 0);

        for _ in 0..(w * h) {
            out.push(b);
            out.push(g);
            out.push(r);
            out.push(a);
        }

        out.resize(out.len() + (and_bytes as usize), 0);
        out
    }

    let out_dir = PathBuf::from(env::var("OUT_DIR").unwrap());
    let ico_path = out_dir.join("app.ico");
    fs::write(&ico_path, build_ico_16_rgba(59, 130, 246, 255)).unwrap();

    let rc_path = out_dir.join("novel-outline-tool.generated.rc");
    let ico_str = ico_path.to_string_lossy().replace('\\', "/");
    let rc = format!("1 24 \"resources/app.manifest\"\n101 ICON \"{}\"\n", ico_str);
    fs::write(&rc_path, rc).unwrap();

    embed_resource::compile(rc_path.to_str().unwrap(), embed_resource::NONE);
    println!("cargo:rerun-if-changed=resources/app.manifest");
}
