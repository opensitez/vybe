use std::io::Write;
use std::process::{Command, Stdio};

#[test]
fn binary_debugger_can_quit_before_guest_trap() {
    // (module (func (export "_start") unreachable))
    let wasm = [
        0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00, 0x01, 0x04, 0x01, 0x60, 0x00, 0x00, 0x03,
        0x02, 0x01, 0x00, 0x07, 0x0a, 0x01, 0x06, b'_', b's', b't', b'a', b'r', b't', 0x00, 0x00,
        0x0a, 0x05, 0x01, 0x03, 0x00, 0x00, 0x0b,
    ];
    let path = std::env::temp_dir().join(format!("vybex-debug-entry-{}.wasm", std::process::id()));
    std::fs::write(&path, wasm).unwrap();
    let mut child = Command::new(env!("CARGO_BIN_EXE_vybex"))
        .arg("--debug")
        .arg(&path)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    child.stdin.take().unwrap().write_all(b"q\n").unwrap();
    let output = child.wait_with_output().unwrap();
    std::fs::remove_file(path).unwrap();
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(output.status.success(), "{stderr}");
    assert!(stderr.contains("Debugger quit"), "{stderr}");
    assert!(!stderr.contains("Runtime error:"), "{stderr}");
}
