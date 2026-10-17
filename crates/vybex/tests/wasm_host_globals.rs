use std::process::{Command, Output};

fn leb(out: &mut Vec<u8>, mut value: usize) {
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        out.push(byte | if value == 0 { 0 } else { 0x80 });
        if value == 0 {
            return;
        }
    }
}

fn name(out: &mut Vec<u8>, value: &str) {
    leb(out, value.len());
    out.extend_from_slice(value.as_bytes());
}

fn section(out: &mut Vec<u8>, id: u8, body: &[u8]) {
    out.push(id);
    leb(out, body.len());
    out.extend_from_slice(body);
}

fn fixture(global: &str) -> Vec<u8> {
    let mut wasm = b"\0asm\x01\0\0\0".to_vec();
    section(
        &mut wasm,
        1,
        &[
            3, 0x60, 1, 0x6f, 1, 0x7f, 0x60, 0, 0, 0x60, 2, 0x6f, 0x6f, 1, 0x7f,
        ],
    );
    let mut imports = vec![4];
    for _ in 0..2 {
        name(&mut imports, "env");
        name(&mut imports, global);
        imports.extend_from_slice(&[3, 0x6f, 0]);
    }
    for (module, function, ty) in [("wasm:js-undefined", "test", 0), ("ecma:object", "is", 2)] {
        name(&mut imports, module);
        name(&mut imports, function);
        imports.extend_from_slice(&[0, ty]);
    }
    section(&mut wasm, 2, &imports);
    section(&mut wasm, 3, &[1, 1]);
    let mut exports = vec![1];
    name(&mut exports, "_start");
    exports.extend_from_slice(&[0, 2]);
    section(&mut wasm, 7, &exports);
    // Both imports must be defined and refer to the same host object.
    let body = [
        0, 0x23, 0, 0x10, 0, 0x04, 0x40, 0x00, 0x0b, 0x23, 1, 0x10, 0, 0x04, 0x40, 0x00, 0x0b,
        0x23, 0, 0x23, 1, 0x10, 1, 0x45, 0x04, 0x40, 0x00, 0x0b, 0x0b,
    ];
    let mut code = vec![1];
    leb(&mut code, body.len());
    code.extend_from_slice(&body);
    section(&mut wasm, 10, &code);
    wasm
}

fn run(global: &str) -> Output {
    let path = std::env::temp_dir().join(format!(
        "vybex-host-global-{global}-{}.wasm",
        std::process::id(),
    ));
    std::fs::write(&path, fixture(global)).unwrap();
    let output = Command::new(env!("CARGO_BIN_EXE_vybex"))
        .arg(&path)
        .output()
        .unwrap();
    std::fs::remove_file(path).unwrap();
    output
}

#[test]
fn registered_host_global_preserves_reference_identity() {
    let output = run("globalThis");
    assert!(
        output.status.success(),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
}

#[test]
fn missing_host_global_fails_linking_before_guest_execution() {
    let output = run("unprovided_host_global");
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(!output.status.success(), "{stderr}");
    assert!(
        stderr.contains("WASM link error: Unresolved global import:"),
        "{stderr}"
    );
    assert!(stderr.contains("unprovided_host_global"), "{stderr}");
    assert!(!stderr.contains("unreachable executed"), "{stderr}");
}
