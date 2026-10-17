#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/byte_length_bulk_before.rs"]
mod before;
#[path = "fixtures/byte_length_bulk_candidate.rs"]
mod candidate;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn validate_native(bytes: &[u8]) {
    // wasmparser does not implement stringref. Validate the complete module
    // with V8's genuine proposal implementation, before native execution.
    use std::io::Write;
    let mut child=std::process::Command::new("node")
        .args(["--experimental-wasm-stringref","-e","const fs=require('node:fs');new WebAssembly.Module(fs.readFileSync(0),{builtins:['js-string'],importedStringConstants:'wasm:string-constants'});"])
        .stdin(std::process::Stdio::piped()).stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::piped()).spawn().unwrap();
    child.stdin.take().unwrap().write_all(bytes).unwrap();
    let out = child.wait_with_output().unwrap();
    assert!(
        out.status.success(),
        "complete stringref module: {}",
        String::from_utf8_lossy(&out.stderr)
    );
}
fn program(fast: bool) -> Vec<Chunk> {
    program_with(if fast {
        primitives::strings::emit_byte_length
    } else {
        before::emit_byte_length
    })
}
fn program_with(emit: fn(&mut [Chunk], usize, u32)) -> Vec<Chunk> {
    let mut root = Chunk::new("byte-length-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = Chunk::new("entry");
    c.arity = 1;
    c.param_count = 1;
    c.local_count = 1;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let mut cs = vec![root, c];
    emit(&mut cs, 1, 0);
    cs[1].emit_op(Op::RETURN, 0);
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn vm() -> VM {
    let mut vm = VM::new();
    primitives::platforms::register_platforms_all(&mut vm);
    vm
}
fn output(name: &str, value: &serde_json::Value) {
    std::fs::write(
        std::path::PathBuf::from(std::env::var("VYBE_PRIMITIVE_ARTIFACT_DIR").unwrap())
            .join(format!("BYTE_LENGTH_{name}.json")),
        serde_json::to_vec_pretty(value).unwrap(),
    )
    .unwrap();
}
#[test]
fn utf8_builtin_rejects_nonstrings_without_coercion() {
    fn direct(chunks: &mut [Chunk], current: usize, line: u32) {
        let c = &mut chunks[current];
        let import = c.add_import("wasm:text-encoder", "measureStringAsUTF8");
        c.emit_call(import, 1, line);
    }
    let mut v = vm();
    let f = v.run(program_with(direct)).unwrap();
    for text in ["", "a\0b", "é界😀"] {
        assert_eq!(
            v.invoke(&f, &[Value::String(text.into())]).unwrap(),
            Value::I32(text.len() as i32)
        );
    }
    for value in [
        Value::Null,
        Value::Undefined,
        Value::I32(1),
        Value::Bool(true),
    ] {
        assert!(v.invoke(&f, &[value]).is_err());
    }
}
#[test]
fn declared_value_abi_caller_remains_valid() {
    for fast in [false, true] {
        let mut cs = program(fast);
        cs[0] = Chunk::new("owner");
        cs[0].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        cs[0].emit_op(Op::RETURN, 0);
        cs[1].func_sig = Some(("externref".into(), "externref".into()));
        native_export::export_chunks_with_validator(
            &cs,
            &format!("byte-length-declared-{fast}"),
            validate_native,
        );
    }
}
#[test]
fn complete_module_sizes_at_multiple_callsites() {
    let mut rows = vec![];
    for sites in [1, 4, 16, 64] {
        let mut sizes = vec![];
        for fast in [false, true] {
            let mut root = Chunk::new("owner");
            primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
            root.emit_op(Op::RETURN, 0);
            let mut entry = Chunk::new("entry");
            entry.arity = 1;
            entry.param_count = 1;
            entry.local_count = 1;
            let mut cs = vec![root, entry];
            for at in 0..sites {
                cs[1].emit_op_u16(Op::LOCAL_GET, 0, 0);
                if fast {
                    primitives::strings::emit_byte_length(&mut cs, 1, 0);
                } else {
                    before::emit_byte_length(&mut cs, 1, 0);
                }
                if at + 1 < sites {
                    cs[1].emit_op(Op::DROP, 0);
                }
            }
            cs[1].emit_op(Op::RETURN, 0);
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                &cs,
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                },
            );
            validate_native(&bytes);
            sizes.push(serde_json::json!({"guest":cs[1].code.len(),"native":bytes.len(),"locals":cs[1].local_count}));
        }
        assert!(sizes[1]["guest"].as_u64() < sizes[0]["guest"].as_u64());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        rows.push(serde_json::json!({"callsites":sites,"sizes":sizes}));
    }
    output("CALLSITE_SIZES", &serde_json::json!(rows));
}
#[test]
fn fallback_counter_wraps_without_gigabyte_allocations() {
    let mut v = vm();
    let f = v
        .run(program_with(candidate::emit_wrapping_fallback))
        .unwrap();
    for s in ["a", "é", "界", "😀", "aé界😀", "\0", "é😀é😀"] {
        assert_eq!(
            v.invoke(&f, &[Value::String(s.into())]).unwrap(),
            Value::I32(i32::MAX.wrapping_add(s.len() as i32))
        );
    }
    native_export::export_chunks_with_validator(
        &program_with(candidate::emit_wrapping_fallback),
        "byte-length-wrapping",
        validate_native,
    );
    let unsigned = v
        .run(program_with(candidate::emit_unsigned_wrapping_fallback))
        .unwrap();
    for s in ["a", "é", "界", "😀", "aé界😀", "\0", "é😀é😀"] {
        assert_eq!(
            v.invoke(&unsigned, &[Value::String(s.into())]).unwrap(),
            Value::I32((-1i32).wrapping_add(s.len() as i32))
        );
    }
    native_export::export_chunks_with_validator(
        &program_with(candidate::emit_unsigned_wrapping_fallback),
        "byte-length-u32-wrapping",
        validate_native,
    );
}
#[test]
fn unicode_values_sizes_and_dispatches() {
    // Compare the actual production emitter with the frozen loop baseline.
    let measured = program_with(candidate::emit_byte_length);
    let production = program(true);
    assert_eq!(production[1].code, measured[1].code);
    assert_eq!(production[1].imports, measured[1].imports);
    assert_eq!(production[1].constants, measured[1].constants);
    assert_eq!(production[1].local_count, measured[1].local_count);
    let mut v = vm();
    let fs = [
        v.run(program(false)).unwrap(),
        v.run(program(true)).unwrap(),
        v.run(program_with(candidate::emit_fallback)).unwrap(),
    ];
    let mut checks = 0;
    let mut check = |s: String| {
        for f in &fs {
            assert_eq!(
                v.invoke(f, &[Value::String(s.clone().into())]).unwrap(),
                Value::I32(s.len() as i32)
            );
            checks += 1;
        }
    };
    for base in (0..0x110000u32).step_by(1024) {
        let s: String = (base..base + 1024).filter_map(char::from_u32).collect();
        check(s);
    }
    for cp in [
        0, 1, 126, 127, 128, 129, 2047, 2048, 2049, 0xd7ff, 0xe000, 65535, 65536, 65537, 0x10ffff,
    ] {
        check(char::from_u32(cp).unwrap().to_string());
    }
    for s in ["", "hello", "a\0b", "é界😀é", "😀😀", "a😀b界"] {
        check(s.into());
    }
    let mut sizes = vec![];
    let mut dispatches = vec![];
    for (fast, f) in [false, true].into_iter().zip(&fs) {
        let cs = program(fast);
        let bytes = vybe_platform_wasm::write_wasm_with_options(
            &cs,
            vybe_platform_wasm::WasmWriteOptions {
                include_name_section: false,
                include_vybe_metadata: false,
            },
        );
        validate_native(&bytes);
        native_export::export_chunks_with_validator(
            &cs,
            &format!("byte-length-{fast}"),
            validate_native,
        );
        sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":bytes.len()}));
        v.record_runtime_perf(true);
        v.invoke(f, &[Value::String("aé界😀".repeat(16).into())])
            .unwrap();
        dispatches.push(v.take_runtime_perf().unwrap().dispatches);
    }
    native_export::export_chunks_with_validator(
        &program_with(candidate::emit_fallback),
        "byte-length-fallback",
        validate_native,
    );
    output(
        "PRE_SIZE_GATE",
        &serde_json::json!({"checks":checks,"sizes":sizes,"dispatches":dispatches}),
    );
    assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
    assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
    assert!(dispatches[1] < dispatches[0]);
    let mut dispatch_cases = vec![];
    for text in ["", "abc", "ééé", "界界界", "😀😀😀", "aé界😀"] {
        let mut counts = vec![];
        for f in &fs {
            v.record_runtime_perf(true);
            v.invoke(f, &[Value::String(text.into())]).unwrap();
            counts.push(v.take_runtime_perf().unwrap().dispatches);
        }

        dispatch_cases.push(serde_json::json!({"text":text,"dispatches":counts}));
    }
    let result = serde_json::json!({"checks":checks,"sizes":sizes,"dispatches":dispatches,"dispatch_cases":dispatch_cases});
    output("GATES", &result);
    eprintln!("{result}");
}
#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, text, control) in [
        ("control", "abcd".repeat(64), true),
        ("empty", String::new(), false),
        ("short", "abc".into(), false),
        ("ascii", "abcd".repeat(64), false),
        ("two-byte", "é".repeat(256), false),
        ("three-byte", "界".repeat(256), false),
        ("astral", "😀".repeat(128), false),
        ("mixed", "aé界😀".repeat(64), false),
    ] {
        let mut v = vm();
        let a = v.run(program(false)).unwrap();
        let b = if control {
            a.clone()
        } else {
            v.run(program(true)).unwrap()
        };
        let fs = [a, b];
        let args = [Value::String(text.clone().into())];
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(v.invoke(f, &args).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        let mut cpu = [vec![], vec![]];
        let mut pair_calls = vec![];
        for pair in 0..7 {
            let mut elapsed = [0., 0.];
            let mut cpus = [0., 0.];
            let mut calls = [0usize, 0];
            let mut turn = pair;
            while elapsed.iter().any(|s| *s < 0.1) {
                for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                    if elapsed[side] >= 0.1 {
                        continue;
                    }
                    let c = thread_cpu_seconds();
                    let start = Instant::now();
                    std::hint::black_box(v.invoke(&fs[side], &args).unwrap());
                    elapsed[side] += start.elapsed().as_secs_f64();
                    cpus[side] += thread_cpu_seconds() - c;
                    calls[side] += 1;
                }
                turn += 1;
            }
            for side in 0..2 {
                samples[side].push(elapsed[side] / calls[side] as f64);
                cpu[side].push(cpus[side] / calls[side] as f64);
            }
            pair_calls.push(calls);
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        let mut cr: Vec<_> = cpu[0].iter().zip(&cpu[1]).map(|(a, b)| a / b).collect();
        cr.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"label":label,"control":control,"units":text.encode_utf16().count(),"median_speedup":ratios[3],"median_thread_cpu_speedup":cr[3],"seconds_per_call":samples,"cpu_per_call":cpu,"pair_calls":pair_calls}));
    }
    let result = serde_json::json!({"same_vm":true,"independent_calls_interleaved":true,"pairs":7,"minimum_sample_ms":100,"cases":rows});
    output("MEASUREMENTS", &result);
    eprintln!("{result}");
}

#[cfg(target_os = "macos")]
fn thread_cpu_seconds() -> f64 {
    #[repr(C)]
    struct Timespec {
        sec: i64,
        nsec: i64,
    }
    unsafe extern "C" {
        fn clock_gettime(clock: i32, result: *mut Timespec) -> i32;
    }
    let mut t = Timespec { sec: 0, nsec: 0 };
    assert_eq!(unsafe { clock_gettime(16, &mut t) }, 0);
    t.sec as f64 + t.nsec as f64 * 1e-9
}
#[cfg(not(target_os = "macos"))]
fn thread_cpu_seconds() -> f64 {
    0.0
}
