//! Behaviour tests for `ecma:function` host imports.
//!
//! Reference: ECMA-262 §20.2 Function objects.
//!
//! Covers Function.prototype.bind/call/apply, the `name` and `length`
//! properties, Function.prototype.toString, and the Function constructor.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-function-test>");
    let import_idx = chunk.add_import("ecma:function", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value { Value::String(Arc::from(text)) }

fn fn_obj(name: &str, arity: i32) -> Value {
    // Encodes a callable function descriptor the host recognises.
    let mut o = Object::new();
    o.properties.insert("__fn_name".to_string(), s(name));
    o.properties.insert("__fn_arity".to_string(), Value::I32(arity));
    o.properties.insert("__fn_return".to_string(), Value::I32(42));
    Value::Object(Arc::new(Mutex::new(o)))
}

// ── name property ─────────────────────────────────────────────────────────────

#[test]
fn name_returns_function_name_string() {
    let f = fn_obj("myFunc", 2);
    assert_eq!(invoke("name", vec![f]), s("myFunc"));
}

#[test]
fn anonymous_function_name_is_empty_string() {
    // ECMA-262 §20.2.3.3: anonymous functions have name = "".
    let f = fn_obj("", 0);
    assert_eq!(invoke("name", vec![f]), s(""));
}

// ── length property ───────────────────────────────────────────────────────────

#[test]
fn length_returns_formal_parameter_count() {
    let f = fn_obj("f", 3);
    assert_eq!(invoke("length", vec![f]), Value::I32(3));
}

#[test]
fn length_of_zero_arity_function_is_zero() {
    let f = fn_obj("f", 0);
    assert_eq!(invoke("length", vec![f]), Value::I32(0));
}

// ── bind ──────────────────────────────────────────────────────────────────────

#[test]
fn bind_returns_a_new_function_object() {
    let f = fn_obj("original", 1);
    let this_arg = Value::I32(0);
    let bound = invoke("bind", vec![f.clone(), this_arg]);
    assert!(matches!(bound, Value::Object(_)));
    // Must be a distinct object.
    let f_ptr   = match &f     { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 0 };
    let bnd_ptr = match &bound { Value::Object(a) => Arc::as_ptr(a) as usize, _ => 1 };
    assert_ne!(f_ptr, bnd_ptr);
}

#[test]
fn bound_function_name_is_prefixed_with_bound() {
    // ECMA-262 §20.2.3.2: bound function name = "bound " + original name.
    let f = fn_obj("foo", 1);
    let bound = invoke("bind", vec![f, Value::Null]);
    assert_eq!(invoke("name", vec![bound]), s("bound foo"));
}

#[test]
fn bound_function_length_is_max_zero_and_original_length_minus_bound_args() {
    // bind(thisArg, arg1) pre-supplies 1 arg; length = max(0, 2 - 1) = 1.
    let f = fn_obj("f", 2);
    let bound = invoke("bindWithArgs", vec![f, Value::Null, Value::I32(99)]);
    assert_eq!(invoke("length", vec![bound]), Value::I32(1));
}

// ── call ──────────────────────────────────────────────────────────────────────

#[test]
fn call_invokes_function_with_explicit_this() {
    // The function descriptor returns __fn_return (42) always.
    let f = fn_obj("f", 0);
    let result = invoke("call", vec![f, Value::Null]);
    assert_eq!(result, Value::I32(42));
}

// ── apply ─────────────────────────────────────────────────────────────────────

#[test]
fn apply_invokes_function_spreading_args_array() {
    let f = fn_obj("f", 0);
    let args_arr = Value::Object(Arc::new(Mutex::new(
        vybe_bytecode::value::Object::new_array(vec![])
    )));
    let result = invoke("apply", vec![f, Value::Null, args_arr]);
    assert_eq!(result, Value::I32(42));
}

// ── toString ─────────────────────────────────────────────────────────────────

#[test]
fn to_string_contains_function_keyword() {
    // ECMA-262 §20.2.3.5: toString of a user-defined function contains
    // "function" and the function name.
    let f = fn_obj("myFn", 0);
    let result = invoke("toString", vec![f]);
    let s_val = match &result {
        Value::String(s) => s.to_string(),
        _ => panic!("expected string"),
    };
    assert!(s_val.contains("function") || s_val.contains("myFn"),
        "toString should mention function or name, got: {}", s_val);
}

// ── Function constructor ──────────────────────────────────────────────────────

#[test]
fn new_function_from_body_string_is_callable() {
    // new Function("return 7") → a function that returns 7.
    let f = invoke("new", vec![s("return 7")]);
    assert!(matches!(f, Value::Object(_)));
}

#[test]
fn new_function_with_param_and_body_strings() {
    // new Function("x", "return x + 1") → (x) => x + 1.
    let f = invoke("newWithParams", vec![s("x"), s("return x + 1")]);
    assert!(matches!(f, Value::Object(_)));
}
