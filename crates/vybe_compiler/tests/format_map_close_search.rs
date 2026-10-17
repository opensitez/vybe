use vybe_compiler::primitives::{bundle, functions, globals, platforms, strings};
use vybe_runtime::value::{Function, Object, ObjectKind};
use vybe_runtime::{Chunk, Op, VM, Value};
#[path = "fixtures/format_map_unit_scan.rs"]
mod baseline;

fn program(old: bool, calls: usize, callback: bool) -> Vec<Chunk> {
    let mut entry = Chunk::new("<format-map-optimization>");
    let body = if old {
        baseline::build_format_map(&mut entry)
    } else {
        strings::build_format_map(&mut entry)
    };
    let mut chunks = vec![entry, body];
    if callback {
        let mut getitem = Chunk::new("<format-map-getitem>");
        getitem.arity = 2;
        getitem.local_count = 2;
        getitem.emit_op_u16(Op::LOCAL_GET, 1, 0);
        let index = getitem.add_import("test:format-map", "getitem");
        getitem.emit_call(index, 1, 0);
        getitem.emit_op(Op::RETURN, 0);
        chunks.push(getitem);
    }
    for call in 0..calls {
        chunks[0].emit_op_u16(Op::REF_FUNC, 1, 0);
        chunks[0].emit(0, 0);
        globals::emit_read(&mut chunks[0], "text", 0);
        globals::emit_read(&mut chunks[0], "dict", 0);
        functions::emit_call(&mut chunks[0], 2, 0);
        if call + 1 != calls {
            chunks[0].emit_op(Op::DROP, 0);
        }
    }
    chunks[0].emit_op(Op::RETURN, 0);
    bundle::finalize_with_runtime_helpers(&mut chunks);
    globals::declare_free_globals(&mut chunks);
    globals::normalize_global_table(&mut chunks);
    chunks
}
pub(crate) fn dict() -> Value {
    let mut object = Object::new();
    for (key, value) in [
        ("a", Value::I32(17)),
        ("b", Value::Bool(true)),
        ("", Value::String("empty".into())),
        ("null", Value::Null),
        ("undef", Value::Undefined),
        ("é", Value::String("café".into())),
    ] {
        object.properties.insert(key.into(), value);
    }
    Value::Object(vybe_runtime::heap::alloc(object))
}
fn vm() -> VM {
    let mut vm = VM::new();
    platforms::register_platforms_all(&mut vm);
    vm
}
fn result(
    runtime: &mut VM,
    chunks: &[Chunk],
    text: Value,
    dictionary: Value,
) -> Result<Value, String> {
    runtime.set_global("text", text);
    runtime.set_global("dict", dictionary);
    runtime.run(chunks.to_vec()).map_err(|mut e| {
        for frame in &mut e.call_stack {
            frame.offset = 0;
        }
        e.to_string()
    })
}
fn validate(chunks: &[Chunk]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&vybe_platform_wasm::write_wasm(chunks))
        .unwrap();
}
#[test]
fn format_map_preserves_legacy_unicode_braces_coercion_and_errors() {
    let old = program(true, 1, false);
    let new = program(false, 1, false);
    validate(&old);
    validate(&new);
    let mut before = vm();
    let mut after = vm();
    for text in [
        "",
        "plain",
        "{a}:{b}:{a}",
        "{missing}:{null}:{undef}",
        "{}",
        "{{a}}",
        "}}x{{",
        "{a",
        "{",
        "{a.b}:{a:.2f}",
        "é😀{é}😀",
        "a\0{a}\n",
    ] {
        assert_eq!(
            result(&mut before, &old, Value::String(text.into()), dict()),
            result(&mut after, &new, Value::String(text.into()), dict()),
            "text={text:?}"
        );
    }
    assert_eq!(
        result(&mut after, &new, Value::String("{a}:{b}".into()), dict()).unwrap(),
        Value::String("17:true".into())
    );
    for text in [
        Value::Null,
        Value::Undefined,
        Value::I32(7),
        Value::Bool(false),
    ] {
        assert_eq!(
            result(&mut before, &old, text.clone(), dict()),
            result(&mut after, &new, text.clone(), dict()),
            "receiver={text:?}"
        );
    }
    for dictionary in [Value::Null, Value::Undefined, Value::I32(7)] {
        assert_eq!(
            result(
                &mut before,
                &old,
                Value::String("{a}".into()),
                dictionary.clone()
            ),
            result(&mut after, &new, Value::String("{a}".into()), dictionary)
        );
    }
    let size = |cs: &[Chunk]| cs.iter().map(|c| c.code.len()).sum::<usize>();
    for calls in [1, 20] {
        let old = program(true, calls, false);
        let new = program(false, calls, false);
        validate(&old);
        validate(&new);
        assert!(size(&new) < size(&old));
        println!(
            "format_map calls={calls} bytecode={}=>{} wasm={}=>{}",
            size(&old),
            size(&new),
            vybe_platform_wasm::write_wasm(&old).len(),
            vybe_platform_wasm::write_wasm(&new).len()
        );
    }
}
#[test]
fn format_map_preserves_getitem_and_string_conversion_order() {
    use std::sync::{Arc, Mutex};
    for old in [true, false] {
        let mut runtime = vm();
        let chunks = program(old, 1, true);
        validate(&chunks);
        let log = Arc::new(Mutex::new(Vec::<String>::new()));
        let observed = log.clone();
        runtime.register_host_fn(
            "test:format-map",
            "getitem",
            Box::new(move |_ctx, args| {
                observed
                    .lock()
                    .unwrap()
                    .push(format!("get:{}", args[0].as_str()));
                Value::String(args[0].as_str().into())
            }),
        );
        let observed = log.clone();
        runtime.register_host_fn(
            "ecma:string",
            "String",
            Box::new(move |_ctx, args| {
                observed
                    .lock()
                    .unwrap()
                    .push(format!("str:{}", args[0].as_str()));
                Value::String(args[0].as_str().into())
            }),
        );
        let mut function = Object::new();
        function.kind = ObjectKind::Function(Function {
            name: Some("getitem".into()),
            arity: 2,
            chunk_index: runtime.chunks.len() + 2,
            upvalues: vec![],
        });
        let mut object = Object::new();
        object.properties.insert(
            "__getitem__".into(),
            Value::Object(vybe_runtime::heap::alloc(function)),
        );
        object.properties.insert("a".into(), Value::Null);
        object.properties.insert("b".into(), Value::Null);
        assert_eq!(
            result(
                &mut runtime,
                &chunks,
                Value::String("x{b}-{a}{b}y".into()),
                Value::Object(vybe_runtime::heap::alloc(object))
            )
            .unwrap(),
            Value::String("xb-aby".into())
        );
        assert_eq!(
            *log.lock().unwrap(),
            ["get:b", "str:b", "get:a", "str:a", "get:b", "str:b"]
        );
    }
}

pub(crate) fn kernel(fast: bool, large: bool) -> Vec<Chunk> {
    let mut root = Chunk::new("format-owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0);
    root.emit(0, 0);
    root.emit_op(Op::RETURN, 0);
    let mut c = if !large {
        if fast {
            strings::build_format_map(&mut root)
        } else {
            baseline::build_format_map(&mut root)
        }
    } else {
        let mut c = Chunk::new("format");
        c.arity = 2;
        c.param_count = 2;
        c.local_count = 2;
        if fast {
            strings::emit_format_map_large_into(&mut c, 0, 1);
        } else {
            baseline::emit_format_map_large_into(&mut c, 0, 1);
        }
        c.emit_op(Op::RETURN, 0);
        c
    };
    c.name = "format".into();
    let mut code = vec![root, c];
    bundle::finalize_with_runtime_helpers(&mut code);
    globals::declare_free_globals(&mut code);
    globals::normalize_global_table(&mut code);
    code
}
fn loaded(fast: bool, large: bool) -> (VM, Value) {
    let mut vm = vm();
    let f = vm.run(kernel(fast, large)).unwrap();
    (vm, f)
}
#[test]
fn generic_and_large_scanners_preserve_unclosed_and_unicode_keys() {
    let key = "é😀".repeat(128);
    let mut d = Object::new();
    d.properties
        .insert(key.as_str().into(), Value::String("answer".into()));
    let d = Value::Object(vybe_runtime::heap::alloc(d));
    let mut rows = vec![];
    for large in [false, true] {
        for text in [
            format!("{{{key}}}"),
            format!("{{{key}"),
            format!("x{{{key}}}😀{{missing}}"),
            "{}{{a}}}}{".to_owned(),
        ] {
            let mut out = vec![];
            for fast in [false, true] {
                let (mut vm, f) = loaded(fast, large);
                out.push(
                    vm.invoke(&f, &[Value::String(text.clone().into()), d.clone()])
                        .map_err(|e| e.message),
                );
            }
            assert_eq!(out[0], out[1]);
        }
        let mut sizes = vec![];
        for fast in [false, true] {
            let code = kernel(fast, large);
            let wasm = vybe_platform_wasm::write_wasm(&code);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&wasm)
                .unwrap();
            sizes.push((code[1].code.len(), wasm.len(), code[1].local_count));
        }
        eprintln!("large={large}, sizes={sizes:?}");
        assert!(sizes[1].0 < sizes[0].0);
        // Record full native size separately: the new standard search import
        // has a module-level cost, while the guest instruction stream shrinks.
        rows.push(serde_json::json!({"large":large,"before":sizes[0],"after":sizes[1]}));
    }
    if let Ok(path) = std::env::var("SCAN_GATES_PATH") {
        std::fs::write(path, serde_json::to_vec_pretty(&rows).unwrap()).unwrap();
    }
}
#[test]
#[ignore = "Seven paired opt0 format-map placeholder search batches"]
fn debug_scans() {
    let mut rows = vec![];
    for keylen in [1, 32, 1024] {
        let key = "a".repeat(keylen);
        let text = format!("{{{key}}}");
        let mut d = Object::new();
        d.properties
            .insert(key.as_str().into(), Value::String("answer".into()));
        let args = [
            Value::String(text.into()),
            Value::Object(vybe_runtime::heap::alloc(d)),
        ];
        let mut workers = [loaded(false, false), loaded(true, false)];
        let mut counts = [1, 1];
        let mut ratios = vec![];
        let mut raw = vec![];
        for pair in 0..7 {
            let mut ns = [0, 0];
            for side in [pair % 2, 1 - pair % 2] {
                loop {
                    let start = std::time::Instant::now();
                    for _ in 0..counts[side] {
                        let w = &mut workers[side];
                        assert_eq!(
                            w.0.invoke(&w.1, &args).unwrap(),
                            Value::String("answer".into())
                        );
                    }
                    ns[side] = start.elapsed().as_nanos();
                    if ns[side] >= 100_000_000 {
                        break;
                    }
                    counts[side] *= 2;
                }
            }
            ratios.push((ns[0] as f64 / counts[0] as f64) / (ns[1] as f64 / counts[1] as f64));
            raw.push(serde_json::json!({"counts":counts,"ns":ns}));
        }
        let mut sorted = ratios.clone();
        sorted.sort_by(f64::total_cmp);
        rows.push(
            serde_json::json!({"key_length":keylen,"median":sorted[3],"ratios":ratios,"raw":raw}),
        );
    }
    if let Ok(path) = std::env::var("SCAN_MEASUREMENTS_PATH") {
        std::fs::write(path, serde_json::to_vec_pretty(&rows).unwrap()).unwrap();
    }
}
#[test]
fn long_placeholder_removes_per_unit_vm_dispatch(){
    let key="a".repeat(1024);let mut d=Object::new();d.properties.insert(key.as_str().into(),Value::String("answer".into()));let args=[Value::String(format!("{{{key}}}").into()),Value::Object(vybe_runtime::heap::alloc(d))];let mut counts=vec![];
    for fast in[false,true]{let(mut vm,f)=loaded(fast,false);vm.record_runtime_perf(true);assert_eq!(vm.invoke(&f,&args).unwrap(),Value::String("answer".into()));counts.push(vm.take_runtime_perf().unwrap().dispatches);}
    assert!(counts[1]*10<counts[0]);
    if let Ok(path)=std::env::var("SCAN_DISPATCHES_PATH"){std::fs::write(path,serde_json::to_vec_pretty(&counts).unwrap()).unwrap();}
}
