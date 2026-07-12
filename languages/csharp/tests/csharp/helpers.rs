use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

#[macro_export]
macro_rules! csharp_cases {
    ($($name:ident => { $src:expr, [$($expected:expr),* $(,)?] };)+) => {
        $(
            #[test]
            fn $name() {
                $crate::helpers::assert_csharp($src, &[$($expected),*]);
            }
        )+
    };
}

/// Run C# source through vybex pipeline: pest grammar → walker → common AST → compiler → VM
pub fn run_csharp(src: &str) -> Vec<String> {
    { static R: std::sync::Once = std::sync::Once::new(); R.call_once(vybe_language_csharp::register); }
    let module = vybe_language_csharp::parse(src).expect("C# parse failed");

    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_csharp::profile_source())
            .expect("Failed to parse C# profile");

    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("C# compile failed");

    if std::env::var("VYBEX_DUMP_CHUNK").ok().as_deref() == Some("<script>") {
        if let Some(chunk) = chunks.first() {
            eprintln!(
                "\n-- chunk 0: {} --\n{}",
                chunk.name,
                vybe_bytecode::debug::disassemble(chunk)
            );
            eprintln!("-- constants --");
            for (ci, cv) in chunk.constants.iter().enumerate() {
                eprintln!("  [{ci}] {cv}");
            }
        }
    }

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            let joined = parts.join(" ");
            let mut sink = out.lock().unwrap();
            if joined.contains('\n') {
                let mut lines: Vec<&str> = joined.split('\n').collect();
                if lines.last().map(|s| s.is_empty()).unwrap_or(false) {
                    lines.pop();
                }
                for line in lines {
                    sink.push(line.to_string());
                }
            } else {
                sink.push(joined);
            }
            Value::Null
        }),
    );
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("C# run failed");
    let result = output.lock().unwrap().clone();
    result
}

pub fn run_csharp_one(src: &str) -> String {
    run_csharp(src).into_iter().next().unwrap_or_default()
}

pub fn assert_csharp(src: &str, expected: &[&str]) {
    let actual = run_csharp(src);
    let expected_vec: Vec<String> = expected.iter().map(|line| (*line).to_string()).collect();
    assert_eq!(actual, expected_vec);
}

pub fn compile_csharp_to_wasm(src: &str) -> Vec<u8> {
    let module = vybe_language_csharp::parse(src).expect("C# parse failed");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_csharp::profile_source())
            .expect("Failed to parse C# profile");
    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("C# compile failed");
    vybe_platform_wasm::write_wasm(&chunks)
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
                let module = std::str::from_utf8(&wasm[cursor..cursor + module_len])
                    .unwrap_or("")
                    .to_string();
                cursor += module_len;
                let (name_len, read) = read_leb(wasm, cursor);
                cursor += read;
                let name = std::str::from_utf8(&wasm[cursor..cursor + name_len])
                    .unwrap_or("")
                    .to_string();
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
