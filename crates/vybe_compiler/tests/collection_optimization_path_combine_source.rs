//! Actual JS parse/profile gate. No parser, prelude, profile or runtime changes.
use vybe_compiler::primitives::platforms;
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/collection_optimization_path_combine_before.rs"]
mod baseline;
#[path = "fixtures/collection_optimization_path_combine_compact.rs"]
mod candidate;
fn ops(c: &Chunk) -> Vec<(usize, Op)> {
    let mut p = 0;
    let mut result = vec![];
    while p < c.code.len() {
        let op = Op::from_code_at(&c.code, p);
        result.push((p, op));
        p += 4 + op.operand_format().size_in(&c.code, p + 4);
    }
    assert_eq!(p, c.code.len());
    result
}
fn slot(c: &Chunk, p: usize) -> u16 {
    u16::from_be_bytes(c.code[p + 4..p + 6].try_into().unwrap())
}
fn relocate(emitted: &mut Chunk, original: &Chunk, imports: &mut Vec<vybe_runtime::chunk::Import>) {
    assert_eq!(
        emitted.constants, original.constants,
        "same String constant pool"
    );
    assert_eq!(
        emitted.global_imports, original.global_imports,
        "same global imports"
    );
    for (p, op) in ops(emitted) {
        if op == Op::CALL {
            let imp = emitted.imports[slot(emitted, p) as usize].clone();
            let index = if let Some(i) = imports.iter().position(|v| v == &imp) {
                i
            } else {
                imports.push(imp);
                imports.len() - 1
            };
            emitted.code[p + 4..p + 6].copy_from_slice(
                &u16::try_from(index)
                    .expect("module import budget")
                    .to_be_bytes(),
            );
        } else if op == Op::GLOBAL_GET {
            let key = u32::from_be_bytes(emitted.code[p + 4..p + 8].try_into().unwrap()) as usize;
            let Value::String(name) = &emitted.constants[key] else {
                panic!("String global key")
            };
            let index = original
                .globals
                .iter()
                .position(|v| v == name.as_ref())
                .expect("existing String global");
            emitted.code[p + 4..p + 8].copy_from_slice(&(index as u32).to_be_bytes());
        }
    }
    emitted.imports.clear();
}
fn substitute(old: &[Chunk], n: u8) -> Vec<Chunk> {
    let index = old
        .iter()
        .position(|c| c.name == "path_subject")
        .expect("actual JS function");
    let original = &old[index];
    let decoded = ops(original);
    assert!(!original.is_async && !original.is_generator);
    assert!(
        original.stack_switch_handlers.is_empty(),
        "no unsupported handler offsets"
    );
    assert!(
        original.user_code_offset.is_none(),
        "subject is not prelude script"
    );
    assert!(
        original
            .local_names
            .iter()
            .all(|v| !v.name.starts_with("<private-helper-ref:")),
        "no pending helper markers"
    );
    let mut found = vec![];
    for (k, &(start, op)) in decoded.iter().enumerate() {
        // Actual accepted production begins by packing the evaluated stack
        // arguments into its private GC buffer, then storing its base slot.
        if op != Op::ARRAY_NEW_FIXED || k + 1 >= decoded.len() {
            continue;
        }
        if slot(original, start) != 0
            || slot(original, start + 2) != n as u16
            || decoded[k + 1].1 != Op::LOCAL_SET
        {
            continue;
        }
        let base = slot(original, decoded[k + 1].0);
        let mut frozen = original.clone();
        frozen.code.clear();
        frozen.lines.clear();
        frozen.local_count = base;
        assert!(
            frozen.imports.is_empty(),
            "finalized non-script imports are empty"
        );
        candidate::emit_combine(&mut frozen, n, original.lines[start]);
        let mut imports = old[0].imports.clone();
        relocate(&mut frozen, original, &mut imports);
        assert_eq!(imports, old[0].imports, "frozen body has no new import");
        let end = start + frozen.code.len();
        if end <= original.code.len() && original.code[start..end] == frozen.code {
            found.push((start, end, base, frozen));
        }
    }
    assert_eq!(
        found.len(),
        1,
        "one exact actual shared-path emitter, not Python adapter or approximated dispatcher"
    );
    let (start, end, base, frozen) = found.remove(0);
    assert!(
        decoded.iter().any(|(p, _)| *p == end),
        "exact end instruction boundary"
    );
    assert_eq!(original.lines[start..end], frozen.lines);
    // The replacement has zero net structural depth and no outward branches.
    // Preserve all prefix/suffix bytes; suffix contains no absolute jump operand.
    assert!(
        decoded
            .iter()
            .filter(|(p, _)| *p >= end)
            .all(|(_, op)| matches!(
                *op,
                Op::LOCAL_GET
                    | Op::LOCAL_SET
                    | Op::NULL
                    | Op::RETURN
                    | Op::STRUCT_GET
                    | Op::CALL
                    | Op::DROP
            )),
        "only verified straight-line return/cleanup suffix"
    );
    assert!(
        original
            .block_i32_results
            .iter()
            .all(|p| *p < start || *p >= end)
    );
    assert!(
        original
            .call_indirect_sigs
            .keys()
            .all(|p| *p < start || *p >= end)
    );
    assert!(
        original
            .call_indirect_canon
            .keys()
            .all(|p| *p < start || *p >= end)
    );
    let mut changed = original.clone();
    changed.code.clear();
    changed.lines.clear();
    changed.local_count = base;
    // Isolated emitter metadata starts empty; source metadata is relocated below.
    changed.block_i32_results.clear();
    changed.call_indirect_sigs.clear();
    changed.call_indirect_canon.clear();
    baseline::emit_combine(&mut changed, n, original.lines[start]);
    assert!(
        changed.block_i32_results.is_empty()
            && changed.call_indirect_sigs.is_empty()
            && changed.call_indirect_canon.is_empty()
    );
    assert!(
        changed.local_count >= frozen.local_count,
        "frozen old scratch is wider"
    );
    assert!(
        decoded
            .iter()
            .filter(|(p, op)| *p < start && matches!(*op, Op::LOCAL_GET | Op::LOCAL_SET))
            .all(|(p, _)| slot(original, *p) < base),
        "all prefix locals lie below fresh helper scratch"
    );
    let mut written = std::collections::HashSet::new();
    for &(p, op) in decoded.iter().filter(|(p, _)| *p >= end) {
        if op == Op::LOCAL_SET {
            written.insert(slot(original, p));
        }
        if op == Op::LOCAL_GET {
            let s = slot(original, p);
            assert!(
                s < base || s >= changed.local_count || written.contains(&s),
                "old helper scratch is not suffix-live"
            );
        }
    }
    let mut result = old.to_vec();
    let mut imports = result[0].imports.clone();
    relocate(&mut changed, original, &mut imports);
    let new_end = start + changed.code.len();
    let rebase = |p: usize| {
        if p < start {
            p
        } else {
            assert!(p >= end);
            new_end + (p - end)
        }
    };
    changed.block_i32_results = original
        .block_i32_results
        .iter()
        .map(|p| rebase(*p))
        .collect();
    changed.call_indirect_sigs = original
        .call_indirect_sigs
        .iter()
        .map(|(p, v)| (rebase(*p), v.clone()))
        .collect();
    changed.call_indirect_canon = original
        .call_indirect_canon
        .iter()
        .map(|(p, v)| (rebase(*p), v.clone()))
        .collect();
    let body = std::mem::take(&mut changed.code);
    let lines = std::mem::take(&mut changed.lines);
    changed.code.extend_from_slice(&original.code[..start]);
    changed.code.extend(body);
    changed.code.extend_from_slice(&original.code[end..]);
    changed.lines.extend_from_slice(&original.lines[..start]);
    changed.lines.extend(lines);
    changed.lines.extend_from_slice(&original.lines[end..]);
    changed.local_count = original.local_count.max(changed.local_count);
    let boundaries = ops(&changed)
        .into_iter()
        .map(|(p, _)| p)
        .collect::<std::collections::HashSet<_>>();
    assert!(
        changed
            .block_i32_results
            .iter()
            .all(|p| boundaries.contains(p))
    );
    assert!(
        changed
            .call_indirect_sigs
            .keys()
            .chain(changed.call_indirect_canon.keys())
            .all(|p| boundaries.contains(p))
    );
    assert_eq!(&changed.code[..start], &original.code[..start]);
    assert_eq!(&changed.code[new_end..], &original.code[end..]);
    assert_eq!(&changed.lines[new_end..], &original.lines[end..]);
    result[index] = changed;
    result[0].imports = imports;
    result
}
fn source(n: usize, parameters: bool, bad: Option<usize>) -> String {
    let values = (0..n)
        .map(|i| {
            if Some(i) == bad {
                "null".to_string()
            } else if i + 1 == n {
                "\"/final-root\"".into()
            } else {
                format!("\"{}\"", "é😀segment/".repeat(32))
            }
        })
        .collect::<Vec<_>>();
    let marked = values
        .iter()
        .enumerate()
        .map(|(i, v)| format!("mark({i},{v})"))
        .collect::<Vec<_>>()
        .join(",");
    let names = (0..n)
        .map(|i| format!("p{i}"))
        .collect::<Vec<_>>()
        .join(",");
    let (params, inside, call) = if parameters {
        (names.clone(), names, marked)
    } else {
        (String::new(), marked, String::new())
    };
    format!(
        "import path from 'node:path';\nvar trace='';\nfunction mark(tag,value){{trace=trace+tag+',';return value;}}\nfunction path_subject({params}){{var result=path.join({inside});return result;}}\nvar answer=path_subject({call});\n"
    )
}
fn parse(src: &str) -> Vec<Chunk> {
    static ONCE: std::sync::Once = std::sync::Once::new();
    ONCE.call_once(vybe_language_js::register);
    let ast = vybe_language_js::parse(src).expect("actual JS parser with prelude");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_js::profile_source()).unwrap();
    vybe_compiler::Compiler::with_profile(profile)
        .compile(&ast)
        .expect("actual JS compiler/profile")
}
fn native_pair(old: &[Chunk], new: &[Chunk], label: &str) {
    let validate = |chunks: &[Chunk]| {
        let b = vybe_platform_wasm::write_wasm(chunks);
        let r = wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&b)
            .map(|_| ())
            .map_err(|e| e.to_string());
        (b, r)
    };
    let (oldbytes, a) = validate(old);
    let (newbytes, b) = validate(new);
    if a.is_err() || b.is_err() {
        let root = "/Users/youness/www/html/vybe/tools/primitive_experiments/path_combine";
        std::fs::write(format!("{root}/source-{label}-old.wasm"), oldbytes).unwrap();
        std::fs::write(format!("{root}/source-{label}-new.wasm"), newbytes).unwrap();
        panic!(
            "strict actual JS Source native gate {label}: production={a:?} frozen={b:?}; full baseline failures hold integration"
        );
    }
}
fn run(chunks: Vec<Chunk>) -> (Result<Value, String>, Option<Value>, Option<Value>) {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    let result = vm.run(chunks).map_err(|mut e| {
        for f in &mut e.call_stack {
            f.offset = 0;
        }
        format!("{e:?}")
    });
    (
        result,
        vm.global("answer").cloned(),
        vm.global("trace").cloned(),
    )
}
#[test]
fn actual_js_path_join_long_literals_and_parameters_preserve_source_abi_native_and_trace() {
    // JS ordinary function calls include one implicit receiver argument. The
    // unchanged source ABI permits 254 explicit parameters, not 255; the root
    // baseline run proved the latter overflows commoncalls' u8 argc. A direct
    // 255-argument path.join literal remains required and fully validated.
    for (n, parameters) in [
        (16, false),
        (16, true),
        (64, false),
        (64, true),
        (255, false),
        (254, true),
    ] {
        let old = parse(&source(n, parameters, None));
        let new = substitute(&old, n as u8);
        native_pair(&old, &new, &format!("{n}-{parameters}"));
        let a = run(old);
        let b = run(new);
        assert_eq!(a, b, "n={n} parameters={parameters}");
        assert!(a.0.is_ok());
        assert_eq!(a.1, Some(Value::String("/final-root".into())));
        let trace = (0..n).map(|i| format!("{i},")).collect::<String>();
        assert_eq!(
            a.2,
            Some(Value::String(trace.into())),
            "positive evaluated-argument order"
        );
    }
}
#[test]
fn actual_js_path_join_invalid_early_argument_keeps_all_argument_evaluation_before_failure() {
    for bad in [0, 8, 14] {
        let n = 16;
        let old = parse(&source(n, true, Some(bad)));
        let new = substitute(&old, n as u8);
        native_pair(&old, &new, &format!("bad-{bad}"));
        let a = run(old);
        let b = run(new);
        assert_eq!(a, b, "bad={bad}");
        assert!(a.0.is_err(), "strict String failure must be observed");
        let trace = (0..n).map(|i| format!("{i},")).collect::<String>();
        assert_eq!(
            a.2,
            Some(Value::String(trace.into())),
            "all caller effects precede first helper type check"
        );
    }
}

#[test]
#[ignore = "exports full real JS Source for separate strict native host-gap diagnostic"]
fn export_actual_js_path_join_native_source_diagnostic() {
    let chunks = parse(&source(16, false, None));
    let bytes = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .expect("full actual JS source strict native validation");
    let dir = std::path::Path::new(file!())
        .parent()
        .unwrap()
        .join("native-source-artifacts");
    std::fs::create_dir_all(&dir).unwrap();
    std::fs::write(dir.join("path-js-source-production.wasm"), bytes).unwrap();
    eprintln!(
        "full JS Source native export only; execution/host coverage remains unproven: {}",
        dir.display()
    );
}
