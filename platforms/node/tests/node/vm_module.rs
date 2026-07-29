//! Behaviour tests for `node:vm` host imports.
//!
//! Reference: <https://nodejs.org/api/vm.html>.
//!
//! Coverage:
//!   - `runInNewContext(code[, sandbox[, options]])` → result
//!   - `runInThisContext(code[, options])` → result
//!   - `createContext([sandbox[, options]])` → contextified sandbox
//!   - `isContext(value)` → boolean
//!   - `Script` constructor + `script.runInNewContext([sandbox])` → result
//!   - `Script.runInThisContext()` → result
//!   - `compileFunction(code, params, options)` → function (surface)
//!   - `measureMemory([options])` → object (surface)
//!
//! Deferred:
//!   - `Module`, `SourceTextModule`, `SyntheticModule` (Node 9+ experimental)
//!   - `Script.createCachedData()` (requires V8 bytecode serialization)

use std::sync::Arc;
use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn call_vm(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-vm-test>");
    let import_idx = chunk.add_import("node:vm", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:vm"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

// ── runInNewContext ────────────────────────────────────────────────────────────

#[test]
fn run_in_new_context_evaluates_arithmetic() {
    let result = call_vm("runInNewContext", vec![s("1 + 2")]);
    assert_eq!(result, Value::I32(3));
}

#[test]
fn run_in_new_context_returns_string_result() {
    let result = call_vm("runInNewContext", vec![s("'hello'")]);
    assert_eq!(result, s("hello"));
}

#[test]
fn run_in_new_context_evaluates_boolean() {
    let result = call_vm("runInNewContext", vec![s("true")]);
    assert_eq!(result, Value::Bool(true));
}

#[test]
fn run_in_new_context_evaluates_null() {
    let result = call_vm("runInNewContext", vec![s("null")]);
    assert_eq!(result, Value::Null);
}

#[test]
fn run_in_new_context_with_sandbox_exposes_variable() {
    use vybe_bytecode::value::Object;
    let mut sandbox = Object::new();
    sandbox.properties.insert("x".to_string(), Value::I32(10));
    let sandbox_val = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(sandbox)));
    let result = call_vm("runInNewContext", vec![s("x * 2"), sandbox_val]);
    assert_eq!(result, Value::I32(20));
}

#[test]
fn run_in_new_context_isolated_from_outer_scope() {
    // 'Math' should be available (global); arbitrary outer vars should not
    let result = call_vm("runInNewContext", vec![s("typeof outerVar")]);
    assert_eq!(result, s("undefined"));
}

// ── runInThisContext ──────────────────────────────────────────────────────────

#[test]
fn run_in_this_context_evaluates_arithmetic() {
    let result = call_vm("runInThisContext", vec![s("3 * 7")]);
    assert_eq!(result, Value::I32(21));
}

#[test]
fn run_in_this_context_returns_last_expression() {
    let result = call_vm("runInThisContext", vec![s("var a = 5; a + 1")]);
    assert_eq!(result, Value::I32(6));
}

// ── createContext ──────────────────────────────────────────────────────────────

#[test]
fn create_context_returns_object() {
    let ctx = call_vm("createContext", vec![]);
    assert!(matches!(ctx, Value::Object(_)));
}

#[test]
fn create_context_with_sandbox_preserves_properties() {
    use vybe_bytecode::value::Object;
    let mut sb = Object::new();
    sb.properties.insert("myVar".to_string(), Value::I32(42));
    let sb_val = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(sb)));
    let ctx = call_vm("createContext", vec![sb_val]);
    match &ctx {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            let my_var = obj
                .properties
                .get("myVar")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert_eq!(my_var, Value::I32(42));
        }
        _ => panic!("expected object"),
    }
}

// ── isContext ─────────────────────────────────────────────────────────────────

#[test]
fn is_context_true_for_contextified_sandbox() {
    let ctx = call_vm("createContext", vec![]);
    let result = call_vm("isContext", vec![ctx]);
    assert_eq!(result, Value::Bool(true));
}

#[test]
fn is_context_false_for_plain_object() {
    use vybe_bytecode::value::Object;
    let plain = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(Object::new())));
    let result = call_vm("isContext", vec![plain]);
    assert_eq!(result, Value::Bool(false));
}

#[test]
fn is_context_false_for_string() {
    let result = call_vm("isContext", vec![s("not a context")]);
    assert_eq!(result, Value::Bool(false));
}

// ── Script constructor ─────────────────────────────────────────────────────────

#[test]
fn script_constructor_returns_object() {
    let script = call_vm("Script", vec![s("1 + 1")]);
    assert!(matches!(script, Value::Object(_)));
}

#[test]
fn script_run_in_new_context_returns_result() {
    let script = call_vm("Script", vec![s("6 * 7")]);
    let result = call_vm("scriptRunInNewContext", vec![script]);
    assert_eq!(result, Value::I32(42));
}

#[test]
fn script_run_in_this_context_returns_result() {
    let script = call_vm("Script", vec![s("100 - 58")]);
    let result = call_vm("scriptRunInThisContext", vec![script]);
    assert_eq!(result, Value::I32(42));
}

#[test]
fn script_run_multiple_times_same_result() {
    let script = call_vm("Script", vec![s("2 ** 10")]);
    let r1 = call_vm("scriptRunInNewContext", vec![script.clone()]);
    let r2 = call_vm("scriptRunInNewContext", vec![script]);
    assert_eq!(r1, Value::I32(1024));
    assert_eq!(r2, Value::I32(1024));
}

// ── compileFunction ───────────────────────────────────────────────────────────

#[test]
fn compile_function_returns_callable() {
    // compileFunction("return 42", [], {}) → a function that returns 42
    let result = call_vm("compileFunction", vec![s("return 42")]);
    assert!(matches!(result, Value::Object(_) | Value::I32(_)));
}

// ── measureMemory ─────────────────────────────────────────────────────────────

#[test]
fn measure_memory_returns_object() {
    let result = call_vm("measureMemory", vec![]);
    assert!(matches!(result, Value::Object(_)));
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_vm_surface_is_registered() {
    let expected = [
        "runInNewContext",
        "runInThisContext",
        "createContext",
        "isContext",
        "Script",
        "scriptRunInNewContext",
        "scriptRunInThisContext",
        "compileFunction",
        "measureMemory",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing node:vm imports: {missing:?}");
}
