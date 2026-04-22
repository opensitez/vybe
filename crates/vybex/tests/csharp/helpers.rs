use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};

/// Run C# source through vybex pipeline: pest grammar → walker → common AST → compiler → VM
pub fn run_csharp(src: &str) -> Vec<String> {
    let module = vybex::languages::csharp::parse(src).expect("C# parse failed");

    let profile = vybex::profile::parse_profile(vybex::languages::csharp::profile_source())
        .expect("Failed to parse C# profile");

    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module).expect("C# compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("C# run failed");
    let result = output.lock().unwrap().clone();
    result
}

pub fn run_csharp_one(src: &str) -> String {
    run_csharp(src).into_iter().next().unwrap_or_default()
}

pub fn compile_csharp_to_wasm(src: &str) -> Vec<u8> {
    let module = vybex::languages::csharp::parse(src).expect("C# parse failed");
    let profile = vybex::profile::parse_profile(vybex::languages::csharp::profile_source())
        .expect("Failed to parse C# profile");
    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("C# compile failed");
    vybe_bytecode::wasm::write_wasm(&chunks)
}

pub fn extract_imports(wasm: &[u8]) -> Vec<(String, String)> {
    let mut out = Vec::new();
    if wasm.len() < 8 || &wasm[..4] != b"\0asm" {
        return out;
    }
    let mut offset = 8;
    while offset < wasm.len() {
        let section_id = wasm[offset];
        offset += 1;
        let (section_size, read) = read_leb(wasm, offset);
        offset += read;
        let end = offset + section_size;
        if section_id == 2 {
            let (count, read) = read_leb(wasm, offset);
            let mut cursor = offset + read;
            for _ in 0..count {
                let (module_len, read) = read_leb(wasm, cursor);
                cursor += read;
                let module = std::str::from_utf8(&wasm[cursor..cursor + module_len]).unwrap_or("").to_string();
                cursor += module_len;
                let (name_len, read) = read_leb(wasm, cursor);
                cursor += read;
                let name = std::str::from_utf8(&wasm[cursor..cursor + name_len]).unwrap_or("").to_string();
                cursor += name_len;
                let kind = wasm[cursor];
                cursor += 1;
                match kind {
                    0 => {
                        let (_, read) = read_leb(wasm, cursor);
                        cursor += read;
                    }
                    1 => {
                        cursor += 1;
                        let (_, read) = read_leb(wasm, cursor);
                        cursor += read;
                        let (_, read) = read_leb(wasm, cursor);
                        cursor += read;
                    }
                    2 => {
                        let flags = wasm[cursor];
                        cursor += 1;
                        let (_, read) = read_leb(wasm, cursor);
                        cursor += read;
                        if flags & 1 != 0 {
                            let (_, read) = read_leb(wasm, cursor);
                            cursor += read;
                        }
                    }
                    3 => cursor += 2,
                    4 => {
                        cursor += 1;
                        let (_, read) = read_leb(wasm, cursor);
                        cursor += read;
                    }
                    _ => break,
                }
                out.push((module, name));
            }
        }
        offset = end;
    }
    out
}

fn read_leb(buf: &[u8], mut offset: usize) -> (usize, usize) {
    let mut value = 0usize;
    let mut shift = 0;
    let start = offset;
    loop {
        let byte = buf[offset];
        offset += 1;
        value |= ((byte & 0x7f) as usize) << shift;
        if byte & 0x80 == 0 {
            break;
        }
        shift += 7;
    }
    (value, offset - start)
}
