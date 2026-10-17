use std::sync::{Arc, Mutex};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::value::{Object, Value};

fn comparison() -> Chunk {
    let mut chunk = Chunk::new("<script>");
    chunk.func_sig = Some(("externref,externref".into(), "externref".into()));
    chunk.arity = 2;
    chunk.local_count = 2;
    chunk.emit_op_u16(Op::LOCAL_GET, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, 1, 0);
    chunk.emit_op(Op::REF_EQ, 0);
    chunk.emit_op(Op::RETURN, 0);
    chunk
}

#[test]
fn writer_reference_comparison_never_calls_abstract_equality() {
    let bytes = vybe_platform_wasm::write_wasm(&[comparison()]);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
    let mut imports = Vec::new();
    let mut calls = Vec::new();
    let mut native_equal = false;
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        match payload.unwrap() {
            wasmparser::Payload::ImportSection(section) => {
                for import in section.into_imports() {
                    let import = import.unwrap();
                    if matches!(import.ty, wasmparser::TypeRef::Func(_)) {
                        imports.push((import.module.to_owned(), import.name.to_owned()));
                    }
                }
            }
            wasmparser::Payload::CodeSectionEntry(body) => {
                for operator in body.get_operators_reader().unwrap() {
                    match operator.unwrap() {
                        wasmparser::Operator::Call { function_index } => calls.push(function_index),
                        wasmparser::Operator::RefEq => native_equal = true,
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }
    assert!(native_equal, "GC references retain native ref.eq");
    assert!(calls.iter().any(|index| imports.get(*index as usize)
        == Some(&("ecma:object".into(), "is".into()))));
    assert!(!calls.iter().any(|index| imports.get(*index as usize)
        == Some(&("ecma:value".into(), "abstractEq".into()))));
}

#[test]
fn host_identity_fallback_agrees_with_vm_for_reference_and_noncoercion_cases() {
    let shared = Arc::new(Mutex::new(Object::new()));
    let distinct = Arc::new(Mutex::new(Object::new()));
    let cases = [
        (Value::Null, Value::Undefined, false),
        (Value::Undefined, Value::Undefined, true),
        (Value::I32(0), Value::String(Arc::from("0")), false),
        (Value::Object(shared.clone()), Value::Object(shared.clone()), true),
        (Value::Object(shared), Value::Object(distinct), false),
    ];
    for (lhs, rhs, expected) in cases {
        let mut vm = VM::new();
        vybe_platform_ecma::register(&mut vm);
        vm.set_global_owned("identity_lhs", lhs.clone());
        vm.set_global_owned("identity_rhs", rhs.clone());
        let mut direct_chunk = Chunk::new("<script>");
        let a = direct_chunk.intern_string_constant("identity_lhs");
        let b = direct_chunk.intern_string_constant("identity_rhs");
        direct_chunk.emit_op_u32(Op::GLOBAL_GET, a, 0);
        direct_chunk.emit_op_u32(Op::GLOBAL_GET, b, 0);
        direct_chunk.emit_op(Op::REF_EQ, 0);
        direct_chunk.emit_op(Op::RETURN, 0);
        let direct = vm.run(vec![direct_chunk]).unwrap().as_i32() != 0;
        let mut host_chunk = Chunk::new("<script>");
        let a = host_chunk.intern_string_constant("identity_lhs");
        let b = host_chunk.intern_string_constant("identity_rhs");
        let import = host_chunk.add_import("ecma:object", "is");
        host_chunk.emit_op_u32(Op::GLOBAL_GET, a, 0);
        host_chunk.emit_op_u32(Op::GLOBAL_GET, b, 0);
        host_chunk.emit_call(import, 2, 0);
        host_chunk.emit_op(Op::RETURN, 0);
        let fallback = vm.run(vec![host_chunk]).unwrap();
        assert_eq!(direct, expected);
        assert_eq!(fallback.as_bool(), expected);
    }
}

#[test]
#[ignore = "requires Node/V8 with the custom-descriptors proposal enabled"]
fn native_v8_host_reference_comparison_is_noncoercing() {
    let path = std::env::temp_dir().join(format!("reference-identity-{}.wasm", std::process::id()));
    std::fs::write(&path, vybe_platform_wasm::write_wasm(&[comparison()])).unwrap();
    let script = r#"
const fs = require('node:fs');
const assert = require('node:assert/strict');
const module_ = new WebAssembly.Module(fs.readFileSync(process.argv[1]));
const shared = {};
let conversions = 0;
const coercible = { valueOf() { conversions++; return 0; } };
for (const [lhs, rhs, expected] of [
  [null, undefined, 0], [undefined, undefined, 1],
  [0, '0', 0], [shared, shared, 1], [shared, {}, 0],
  [coercible, 0, 0],
]) {
  const imports = {};
  for (const entry of WebAssembly.Module.imports(module_)) {
    const namespace = imports[entry.module] ||= {};
    if (entry.kind === 'global') {
      if (entry.module === 'wasm:js-undefined') namespace[entry.name] = undefined;
      else if (entry.module === 'wasm:js-boolean') namespace[entry.name] = entry.name === 'true';
      else throw new Error('Unexpected global: ' + entry.module + '.' + entry.name);
    } else if (entry.kind === 'function') {
      if (entry.module === 'ecma:object' && entry.name === 'is') namespace[entry.name] = Object.is;
      else if (entry.module === 'wasm:js-boolean' && entry.name === 'cast') namespace[entry.name] = x => {
        assert.equal(typeof x, 'boolean'); return x ? 1 : 0;
      };
      else if (entry.module === 'wasm:js-number' && entry.name === 'fromI32') namespace[entry.name] = x => x;
      else namespace[entry.name] = () => { throw new Error('Unexpected import: ' + entry.module + '.' + entry.name); };
    } else throw new Error('Unexpected import kind: ' + entry.kind);
  }
  const instance = new WebAssembly.Instance(module_, imports);
  assert.equal(instance.exports._start(lhs, rhs), expected);
}
assert.equal(conversions, 0);
"#;
    let result = std::process::Command::new("node")
        .args(["--experimental-wasm-custom-descriptors", "-e", script])
        .arg(&path).output();
    std::fs::remove_file(&path).unwrap();
    let output = result.expect("Node must be installed for this explicit V8 gate");
    assert!(output.status.success(), "{}", String::from_utf8_lossy(&output.stderr));
}
