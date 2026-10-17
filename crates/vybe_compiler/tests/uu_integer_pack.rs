#![cfg(feature = "primitive-benchmarks")]
use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/uu_pack_before.rs"]
mod before;
#[path = "fixtures/uu_pack_candidate.rs"]
mod candidate;
#[path = "fixtures/primitive_native_export.rs"]
mod native_export;
fn program_mode(mode: u8, decode: bool) -> Vec<Chunk> {
    let mut c = if mode == 2 {
        let mut wrapper = Chunk::new("uu-wrapper");
        wrapper.arity = 1;
        wrapper.local_count = 1;
        wrapper.emit_op_u16(Op::LOCAL_GET, 0, 0);
        let mut cs = vec![wrapper];
        if decode {
            primitives::string_encoding::emit_convert_uudecode(&mut cs, 0, 1, 0)
        } else {
            primitives::string_encoding::emit_convert_uuencode(&mut cs, 0, 1, 0)
        }
        cs[0].emit_op(Op::RETURN, 0);
        primitives::bundle::finalize_with_runtime_helpers(&mut cs);
        cs.into_iter()
            .find(|c| {
                c.name
                    == if decode {
                        "__stdlib_encoding_convert_uudecode"
                    } else {
                        "__stdlib_encoding_convert_uuencode"
                    }
            })
            .unwrap()
    } else {
        let mut c = Chunk::new("entry");
        c.arity = 1;
        c.local_count = 1;
        c.emit_op_u16(Op::LOCAL_GET, 0, 0);
        let mut cs = vec![c];
        match (mode, decode) {
            (0, false) => before::emit_convert_uuencode(&mut cs, 0, 1, 0),
            (0, true) => before::emit_convert_uudecode(&mut cs, 0, 1, 0),
            (_, false) => candidate::emit_convert_uuencode(&mut cs, 0, 1, 0),
            (_, true) => candidate::emit_convert_uudecode(&mut cs, 0, 1, 0),
        }
        cs[0].emit_op(Op::RETURN, 0);
        cs.pop().unwrap()
    };
    c.name = "entry".into();
    c.param_count = 1;
    let mut root = Chunk::new("uu-pack-owner");
    primitives::functions::emit_ref_func(&mut root, 1, 0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut cs = vec![root, c];
    primitives::globals::declare_free_globals(&mut cs);
    primitives::globals::normalize_global_table(&mut cs);
    cs
}
fn program(fast: bool, decode: bool) -> Vec<Chunk> {
    program_mode(if fast { 2 } else { 0 }, decode)
}
fn vm() -> VM {
    let mut v = VM::new();
    primitives::platforms::register_platforms_all(&mut v);
    v
}
fn function_index(f: &Value) -> usize {
    let Value::Object(o) = f else {
        panic!("function")
    };
    let o = o.lock().unwrap();
    let vybe_runtime::value::ObjectKind::Function(f) = &o.kind else {
        panic!("function")
    };
    f.chunk_index
}
fn enc(s: &str) -> String {
    let us: Vec<u16> = s.encode_utf16().collect();
    let mut out = vec![(us.len() as u32 + 32) as u16];
    for group in us.chunks(3) {
        let a = group[0] as u32;
        let b = *group.get(1).unwrap_or(&0) as u32;
        let c = *group.get(2).unwrap_or(&0) as u32;
        for n in [a / 4, (a % 4) * 16 + b / 16, (b % 16) * 4 + c / 64, c % 64] {
            out.push(if n == 0 { 96 } else { (n + 32) as u16 });
        }
    }
    out.push(10);
    String::from_utf16_lossy(&out)
}
fn dec(s: &str) -> Value {
    let us: Vec<u16> = s.encode_utf16().collect();
    if us.is_empty() {
        return Value::Bool(false);
    }
    let n = us[0].wrapping_sub(32) & 63;
    let needed = 1 + usize::from(n).div_ceil(3) * 4;
    if us.len() < needed {
        return Value::Bool(false);
    }
    let mut out = vec![];
    for group in us[1..needed].chunks(4) {
        let a = group[0].wrapping_sub(32) & 63;
        let b = group[1].wrapping_sub(32) & 63;
        let c = group[2].wrapping_sub(32) & 63;
        let d = group[3].wrapping_sub(32) & 63;
        for x in [a * 4 + b / 16, (b % 16) * 16 + c / 4, (c % 4) * 64 + d] {
            if out.len() < n as usize {
                out.push(x)
            }
        }
    }
    Value::String(String::from_utf16_lossy(&out).into())
}
fn output(n: &str, x: &serde_json::Value) {
    std::fs::write(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("../../UU_PACK_{n}.json")),
        serde_json::to_vec_pretty(x).unwrap(),
    )
    .unwrap();
}
#[test]
fn production_matches_reference() {
    for decode in [false, true] {
        let a = program_mode(2, decode);
        let b = program_mode(1, decode);
        assert_eq!(a[1].code, b[1].code);
        assert_eq!(a[1].constants, b[1].constants);
        assert_eq!(a[1].imports, b[1].imports);
        assert_eq!(a[1].global_imports, b[1].global_imports);
        assert_eq!(a[1].local_count, b[1].local_count);
    }
}
#[test]
fn utf16_packing_bounds_malformed_and_sizes() {
    let mut encode = vec![];
    for cp in (0..=0xffffu32)
        .step_by(17)
        .chain(0..256)
        .chain([0xd7ff, 0xe000, 0xffff, 0x10000, 0x10ffff])
    {
        if let Some(c) = char::from_u32(cp) {
            let s = c.to_string();
            encode.extend([
                s.clone(),
                format!("A{s}"),
                format!("{s}B"),
                format!("AB{s}"),
            ]);
        }
    }
    for n in 0..100 {
        encode.push(
            (0..n)
                .map(|i| char::from_u32((i * 37 % 256) as u32).unwrap())
                .collect(),
        );
    }
    encode.extend(
        ["Cat", "Ω", "é", "😀", "\0\0\0", "ABC", "ÿÿÿ"]
            .iter()
            .map(|s| s.to_string()),
    );
    let mut decode = vec![
        "".into(),
        "`".into(),
        " ".into(),
        "#0V%".into(),
        "#0V%T\n".into(),
        "!80``\n".into(),
    ];
    for cp in (0..=0xffffu32).step_by(31) {
        if let Some(c) = char::from_u32(cp) {
            for at in 0..4 {
                let mut s = String::from("#");
                for k in 0..4 {
                    s.push(if k == at { c } else { 'A' });
                }
                decode.push(s);
            }
        }
    }
    for n in 0..100 {
        let s: String = (0..n)
            .map(|i| char::from_u32((i * 41 % 256) as u32).unwrap())
            .collect();
        let encoded = enc(&s);
        for cut in [
            0,
            1,
            encoded.len() / 2,
            encoded.len().saturating_sub(1),
            encoded.len(),
        ] {
            decode.push(encoded.chars().take(cut).collect());
        }
    }
    let mut checks = 0;
    let mut rows = vec![];
    for decoding in [false, true] {
        let mut v = vm();
        let fs = [
            v.run(program(false, decoding)).unwrap(),
            v.run(program(true, decoding)).unwrap(),
        ];
        let indices = [function_index(&fs[0]), function_index(&fs[1])];
        assert_ne!(indices[0], indices[1]);
        for s in if decoding { &decode } else { &encode } {
            let expected = if decoding {
                dec(s)
            } else {
                Value::String(enc(s).into())
            };
            for f in &fs {
                assert_eq!(
                    v.invoke(f, &[Value::String(s.clone().into())]).unwrap(),
                    expected,
                    "decode={decoding} input={s:?}"
                );
                checks += 1;
            }
        }
        for input in [
            Value::Null,
            Value::Undefined,
            Value::Bool(false),
            Value::I32(7),
        ] {
            let mut outcomes = vec![];
            for fast in [false, true] {
                let mut bad = vm();
                let f = bad.run(program(fast, decoding)).unwrap();
                outcomes.push(bad.invoke(&f, &[input.clone()]).map_err(|e| e.message));
            }
            assert_eq!(outcomes[0], outcomes[1]);
            assert!(outcomes[0].is_err());
            checks += 1;
        }
        let mut sizes = vec![];
        for fast in [false, true] {
            let cs = program(fast, decoding);
            let b = vybe_platform_wasm::write_wasm_with_options(
                &cs,
                vybe_platform_wasm::WasmWriteOptions {
                    include_name_section: false,
                    include_vybe_metadata: false,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&b)
                .unwrap();
            native_export::export_chunks(&cs, &format!("uu-pack-{decoding}-{fast}"));
            sizes.push(serde_json::json!({"body":cs[1].code.len(),"locals":cs[1].local_count,"native":b.len()}));
        }
        assert!(sizes[1]["body"].as_u64() < sizes[0]["body"].as_u64());
        assert!(sizes[1]["native"].as_u64() < sizes[0]["native"].as_u64());
        let mut work = vec![];
        for n in [0, 1, 3, 45, 63, 1024] {
            let s = "a".repeat(n);
            let input = if decoding { enc(&s) } else { s };
            let mut counts = vec![];
            for f in &fs {
                v.record_runtime_perf(true);
                v.invoke(f, &[Value::String(input.clone().into())]).unwrap();
                counts.push(v.take_runtime_perf().unwrap().dispatches);
            }
            work.push(serde_json::json!({"units":n,"dispatches":counts}));
        }
        rows.push(serde_json::json!({"decode":decoding,"chunk_indices":indices,"sizes":sizes,"work":work}));
    }
    output(
        "GATES",
        &serde_json::json!({"checks":checks,"encode_inputs":encode.len(),"decode_inputs":decode.len(),"rows":rows}),
    );
}
fn parsed_php(body: &str) -> Vec<Chunk> {
    static REGISTER: std::sync::Once = std::sync::Once::new();
    REGISTER.call_once(vybe_language_php::register);
    let module = vybe_language_php::parse(&format!("<?php {body}")).unwrap();
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_php::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap()
}
fn frozen_source(chunks: &[Chunk]) -> Vec<Chunk> {
    let mut result = chunks.to_vec();
    for original in chunks {
        let decoding = match original.name.as_str() {
            "__stdlib_encoding_convert_uuencode" => false,
            "__stdlib_encoding_convert_uudecode" => true,
            _ => continue,
        };
        let standalone = program_mode(0, decoding);
        let old = &standalone[1];
        let mut code = old.code.clone();
        let mut ip = 0;
        while ip + 3 < code.len() {
            let op = Op::from_code_at(&code, ip);
            let operand = ip + 4;
            if op == Op::CALL {
                let index =
                    u16::from_be_bytes(code[operand..operand + 2].try_into().unwrap()) as usize;
                let import = &old.imports[index];
                let new = chunks[0]
                    .imports
                    .iter()
                    .position(|v| v.module == import.module && v.name == import.name)
                    .expect("same host import");
                code[operand..operand + 2].copy_from_slice(&(new as u16).to_be_bytes());
            } else if op == Op::GLOBAL_GET || op == Op::GLOBAL_SET {
                let index =
                    u32::from_be_bytes(code[operand..operand + 4].try_into().unwrap()) as usize;
                let name = &old.globals[index];
                let new = original
                    .globals
                    .iter()
                    .position(|v| v == name)
                    .expect("same global name");
                code[operand..operand + 4].copy_from_slice(&(new as u32).to_be_bytes());
            }
            ip = operand + op.operand_format().size_in(&code, operand);
        }
        let at = result.iter().position(|v| v.name == original.name).unwrap();
        result[at].code = code;
        result[at].lines = old.lines.clone();
        result[at].constants = old.constants.clone();
        result[at].local_count = old.local_count;
        result[at].block_i32_results = old.block_i32_results.clone();
    }
    result
}
#[test]
fn parsed_php_coercion_aliases_and_stringable_order() {
    let mut rows = vec![];
    for (body, expected) in [
        (
            "$s='Cat'; return convert_uuencode($s);",
            Value::String(enc("Cat").into()),
        ),
        (
            "return convert_uudecode('#0V%T');",
            Value::String("Cat".into()),
        ),
        (
            "$s='Cat'; return convert_uudecode(convert_uuencode($s));",
            Value::String("Cat".into()),
        ),
        ("return convert_uudecode('`');", Value::String("".into())),
        ("return convert_uudecode('');", Value::Bool(false)),
        (
            "return convert_uuencode(1234);",
            Value::String(enc("1234").into()),
        ),
        (
            "$n=0; class S { public function __construct() {} public function __toString() { global $n; $n++; return 'Cat'; } } $r=convert_uuencode(new S()); return $n.'|'.$r;",
            Value::String(format!("1|{}", enc("Cat")).into()),
        ),
        (
            "class T { public function __construct() {} public function __toString() { throw new Exception('conversion'); } } try { convert_uuencode(new T()); } catch (Exception $e) { return 'caught'; } return 'missed';",
            Value::String("caught".into()),
        ),
    ] {
        let chunks = parsed_php(body);
        for (old, cs) in [(true, frozen_source(&chunks)), (false, chunks)] {
            let mut v = vm();
            let result = v
                .run(cs)
                .unwrap_or_else(|e| panic!("old={old} body={body}: {e:?}"));
            assert_eq!(result, expected, "old={old} body={body}");
        }
        rows.push(serde_json::json!({"source":body,"before_after_expected":true}));
    }
    output(
        "PHP_SOURCE_GATES",
        &serde_json::json!({"cases":rows,"frozen_helper_relocation":"Host imports map to original root import indices; globals map by exact original name; caller/class code unchanged."}),
    );
}

#[test]
fn implicit_stringable_constructor_matches_frozen_helpers() {
    let body = "$n=0; class S { public function __toString() { global $n; $n++; return 'Cat'; } } $r=convert_uuencode(new S()); return $n.'|'.$r;";
    let cs = parsed_php(body);
    let old = vm().run(frozen_source(&cs)).map_err(|e| e.message);
    let new = vm().run(cs).map_err(|e| e.message);
    assert_eq!(old, new, "implicit constructor must not regress");
    if let Ok(value) = &old {
        assert_eq!(value, &Value::String(format!("1|{}", enc("Cat")).into()));
    }
    output(
        "IMPLICIT_STRINGABLE",
        &serde_json::json!({"source":body,"same_before_after":true,"baseline_error":old.err(),"scope":"Frozen pre-change UU bodies relocated into the actual parsed module; all caller/class code unchanged."}),
    );
}

#[test]
fn dynamic_builtin_alias_matches_frozen_helpers() {
    let body = "$f='convert_uuencode'; return $f('Cat');";
    let cs = parsed_php(body);
    let inventory: Vec<_> = cs
        .iter()
        .map(|c| serde_json::json!({"name":c.name,"body":c.code.len(),"locals":c.local_count}))
        .collect();
    let codec_helpers: Vec<_> = cs
        .iter()
        .filter(|c| c.name.contains("encoding_convert_uu"))
        .map(|c| c.name.clone())
        .collect();
    let old = vm().run(frozen_source(&cs)).map_err(|e| e.message);
    let new = vm().run(cs).map_err(|e| e.message);
    assert_eq!(old, new, "dynamic builtin alias must not regress");
    if let Ok(value) = &old {
        assert_eq!(value, &Value::String(enc("Cat").into()));
    }
    output(
        "ALIAS_BASELINE_GAP",
        &serde_json::json!({"source":body,"chunks":inventory,"codec_helpers":codec_helpers,"same_before_after":true,"baseline_error":old.err(),"scope":"Direct Compiler parsed-source fixture. Frozen pre-change UU bodies relocated into the actual parsed module; all caller/class code unchanged. No CLI claim."}),
    );
}

#[test]
#[ignore]
fn debug_pairs() {
    use std::time::Instant;
    let mut rows = vec![];
    for (label, decode, n, control) in [
        ("control", false, 256, true),
        ("encode-empty", false, 0, false),
        ("encode-one", false, 1, false),
        ("encode-short", false, 3, false),
        ("encode-256", false, 256, false),
        ("encode-4096", false, 4096, false),
        ("decode-empty", true, 0, false),
        ("decode-short", true, 3, false),
        ("decode-45", true, 45, false),
        ("decode-63", true, 63, false),
    ] {
        let mut v = vm();
        let a = v.run(program(false, decode)).unwrap();
        let b = if control {
            a.clone()
        } else {
            v.run(program(true, decode)).unwrap()
        };
        let fs = [a, b];
        if !control {
            assert_ne!(function_index(&fs[0]), function_index(&fs[1]));
        }
        let text = (0..n)
            .map(|i| char::from_u32((i * 37 % 256) as u32).unwrap())
            .collect::<String>();
        let text = if decode { enc(&text) } else { text };
        let args = [Value::String(text.into())];
        for f in &fs {
            for _ in 0..8 {
                std::hint::black_box(v.invoke(f, &args).unwrap());
            }
        }
        let mut samples = [vec![], vec![]];
        let mut cpu_samples = [vec![], vec![]];
        for pair in 0..7 {
            let mut secs = [0.0, 0.0];
            let mut cpu_secs = [0.0, 0.0];
            let mut calls = 0;
            let mut turn = pair;
            while secs[0] < 0.5 || secs[1] < 0.5 || calls < 8 {
                for side in if turn % 2 == 0 { [0, 1] } else { [1, 0] } {
                    let cpu = thread_cpu_seconds();
                    let t = Instant::now();
                    std::hint::black_box(v.invoke(&fs[side], &args).unwrap());
                    secs[side] += t.elapsed().as_secs_f64();
                    if let (Some(a), Some(b)) = (cpu, thread_cpu_seconds()) {
                        cpu_secs[side] += b - a;
                    }
                }
                calls += 1;
                turn += 1;
            }
            for side in 0..2 {
                samples[side].push(secs[side] / calls as f64);
                cpu_samples[side].push(cpu_secs[side] / calls as f64);
            }
        }
        let mut ratios: Vec<_> = samples[0]
            .iter()
            .zip(&samples[1])
            .map(|(a, b)| a / b)
            .collect();
        ratios.sort_by(f64::total_cmp);
        let mut cpu_ratios: Vec<_> = cpu_samples[0]
            .iter()
            .zip(&cpu_samples[1])
            .filter(|(_, b)| **b > 0.0)
            .map(|(a, b)| a / b)
            .collect();
        cpu_ratios.sort_by(f64::total_cmp);
        rows.push(serde_json::json!({"label":label,"control":control,"median_speedup":ratios[3],"median_thread_cpu_speedup":cpu_ratios.get(cpu_ratios.len()/2),"seconds_per_call":samples,"thread_cpu_seconds_per_call":cpu_samples}));
    }
    output(
        "MEASUREMENTS",
        &serde_json::json!({"same_vm":true,"pairs":7,"equal_calls_interleaved":true,"minimum_sample_ms":500,"cases":rows}),
    );
}

// Diagnostic only: the local macOS SDK defines CLOCK_THREAD_CPUTIME_ID=16.
// Wall time remains recorded; thread CPU time excludes descheduling by other jobs.
#[cfg(target_os = "macos")]
fn thread_cpu_seconds() -> Option<f64> {
    #[repr(C)]
    struct Timespec {
        sec: i64,
        nsec: i64,
    }
    unsafe extern "C" {
        fn clock_gettime(clock: i32, result: *mut Timespec) -> i32;
    }
    let mut result = Timespec { sec: 0, nsec: 0 };
    assert_eq!(unsafe { clock_gettime(16, &mut result) }, 0);
    Some(result.sec as f64 + result.nsec as f64 * 1e-9)
}
#[cfg(not(target_os = "macos"))]
fn thread_cpu_seconds() -> Option<f64> {
    None
}
