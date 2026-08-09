//! Behaviour tests for `ecma:error` host imports.
//!
//! Reference: ECMA-262 §20.5 Error objects.
//!
//! Each test covers a distinct behaviour.

use std::sync::Arc;
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::Value;
use vybe_runtime::{Chunk, Op, VM};

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn push_arg(vm: &mut VM, chunk: &mut Chunk, value: Value) {
    match value {
        Value::I32(n) => chunk.emit_i32_const(n, 0),
        Value::I64(n) => chunk.emit_i64_const(n, 0),
        Value::F32(f) => chunk.emit_f32_const(f, 0),
        Value::F64(f) => chunk.emit_f64_const(f, 0),
        Value::Bool(b) => chunk.emit_bool_const(b, 0),
        Value::String(s) => chunk.emit_string_const(&s, 0),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let global = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
            );
            vm.globals.insert(global.clone(), other);
            let ci = chunk.intern_string_constant(&global);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
        }
    }
}

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-error-test>");
    let import_idx = chunk.add_import("ecma:error", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

fn prop(obj: &Value, key: &str) -> Value {
    match obj {
        Value::Object(o) => o
            .lock()
            .unwrap()
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Undefined),
        _ => Value::Undefined,
    }
}

// ── Error — message and name ──────────────────────────────────────────────────

#[test]
fn error_message_property_matches_constructor_argument() {
    let e = invoke("Error", vec![s("something went wrong")]);
    assert_eq!(prop(&e, "message"), s("something went wrong"));
}

#[test]
fn error_name_property_is_error_not_subtype() {
    let e = invoke("Error", vec![s("msg")]);
    assert_eq!(prop(&e, "name"), s("Error"));
}

#[test]
fn error_with_no_argument_has_empty_or_undefined_message() {
    let e = invoke("Error", vec![]);
    assert!(
        matches!(prop(&e, "message"), Value::String(ref s) if s.is_empty())
            || matches!(prop(&e, "message"), Value::Undefined),
        "no-arg Error must have empty or undefined message"
    );
}

#[test]
fn error_has_stack_property() {
    let e = invoke("Error", vec![s("trace me")]);
    assert!(
        !matches!(prop(&e, "stack"), Value::Undefined),
        "Error.stack must be present"
    );
}

// ── Each subtype has its own name ─────────────────────────────────────────────

#[test]
fn eval_error_name_is_eval_error() {
    let e = invoke("EvalError", vec![s("bad eval")]);
    assert_eq!(prop(&e, "name"), s("EvalError"));
}

#[test]
fn range_error_name_is_range_error() {
    let e = invoke("RangeError", vec![s("out of range")]);
    assert_eq!(prop(&e, "name"), s("RangeError"));
}

#[test]
fn reference_error_name_is_reference_error() {
    let e = invoke("ReferenceError", vec![s("undefined variable")]);
    assert_eq!(prop(&e, "name"), s("ReferenceError"));
}

#[test]
fn syntax_error_name_is_syntax_error() {
    let e = invoke("SyntaxError", vec![s("unexpected token")]);
    assert_eq!(prop(&e, "name"), s("SyntaxError"));
}

#[test]
fn type_error_name_is_type_error() {
    let e = invoke("TypeError", vec![s("not a function")]);
    assert_eq!(prop(&e, "name"), s("TypeError"));
}

#[test]
fn uri_error_name_is_uri_error() {
    let e = invoke("URIError", vec![s("malformed URI")]);
    assert_eq!(prop(&e, "name"), s("URIError"));
}

// ── Subtype message is preserved independently of name ────────────────────────

#[test]
fn type_error_message_is_independent_of_its_name() {
    // Confirm name≠message; a common implementation mistake is to set both the same.
    let e = invoke("TypeError", vec![s("cannot read property")]);
    assert_eq!(prop(&e, "message"), s("cannot read property"));
    assert_ne!(prop(&e, "name"), prop(&e, "message"));
}

// ── Two error instances are distinct objects ───────────────────────────────────

#[test]
fn two_error_instances_are_not_the_same_object() {
    let e1 = invoke("Error", vec![s("a")]);
    let e2 = invoke("Error", vec![s("b")]);
    let p1 = match &e1 {
        Value::Object(a) => std::sync::Arc::as_ptr(a) as usize,
        _ => 0,
    };
    let p2 = match &e2 {
        Value::Object(a) => std::sync::Arc::as_ptr(a) as usize,
        _ => 1,
    };
    assert_ne!(p1, p2);
}

// ── AggregateError ────────────────────────────────────────────────────────────

#[test]
fn aggregate_error_name_is_aggregate_error() {
    use std::sync::Mutex;
    use vybe_runtime::value::Object;
    let errors = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![]))));
    let e = invoke("AggregateError", vec![errors]);
    assert_eq!(prop(&e, "name"), s("AggregateError"));
}

#[test]
fn aggregate_error_has_errors_property_as_array() {
    use std::sync::Mutex;
    use vybe_runtime::value::{Object, ObjectKind};
    let errors = Value::Object(Arc::new(Mutex::new(Object::new_array(vec![
        invoke("TypeError", vec![s("first")]),
        invoke("RangeError", vec![s("second")]),
    ]))));
    let e = invoke("AggregateError", vec![errors]);
    match prop(&e, "errors") {
        Value::Object(o) => assert!(matches!(o.lock().unwrap().kind, ObjectKind::Array(_))),
        other => panic!("errors must be array, got {:?}", other),
    }
}

// ── Error.cause (ES2022 §20.5.5.1) ───────────────────────────────────────────

#[test]
fn error_cause_is_preserved_from_options_object() {
    // ECMA-262 ES2022: new Error("msg", { cause: originalError }) stores cause.
    use std::sync::Mutex;
    use vybe_runtime::value::Object;
    let cause = invoke("TypeError", vec![s("root cause")]);
    let opts = {
        let mut o = Object::new();
        o.properties.insert("cause".to_string(), cause.clone());
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let e = invoke("ErrorWithCause", vec![s("wrapper"), opts]);
    assert!(matches!(prop(&e, "cause"), Value::Object(_)));
}

#[test]
fn error_without_cause_option_has_no_cause_property() {
    // When no options or no cause key, .cause must be Undefined.
    let e = invoke("Error", vec![s("plain error")]);
    assert_eq!(prop(&e, "cause"), Value::Undefined);
}

// ── Error instanceof checks via __type ────────────────────────────────────────

#[test]
fn each_error_subtype_has_distinct_name_from_generic_error() {
    // All subtypes must have a name that differs from "Error".
    for kind in &[
        "TypeError",
        "RangeError",
        "SyntaxError",
        "ReferenceError",
        "URIError",
        "EvalError",
    ] {
        let e = invoke(kind, vec![s("test")]);
        assert_ne!(
            prop(&e, "name"),
            s("Error"),
            "{kind} must not have name 'Error'"
        );
    }
}

// ── Error.isError (ES2025 §20.5.2.1) ─────────────────────────────────────────

#[test]
fn is_error_true_for_generic_error_instance() {
    // ECMA-262 ES2025: Error.isError(e) returns true for any Error instance.
    let e = invoke("Error", vec![s("test")]);
    assert_eq!(invoke("isError", vec![e]), Value::Bool(true));
}

#[test]
fn is_error_true_for_subtype_error_instances() {
    for kind in &["TypeError", "RangeError", "SyntaxError"] {
        let e = invoke(kind, vec![s("test")]);
        let result = invoke("isError", vec![e]);
        assert_eq!(
            result,
            Value::Bool(true),
            "isError must return true for {kind}"
        );
    }
}

#[test]
fn is_error_false_for_plain_objects() {
    // Error.isError({}) → false; plain objects are not Errors.
    use std::sync::Mutex;
    use vybe_runtime::value::Object;
    let obj = Value::Object(Arc::new(Mutex::new(Object::new())));
    assert_eq!(invoke("isError", vec![obj]), Value::Bool(false));
}

#[test]
fn is_error_false_for_primitives() {
    assert_eq!(
        invoke("isError", vec![s("error string")]),
        Value::Bool(false)
    );
    assert_eq!(invoke("isError", vec![Value::I32(42)]), Value::Bool(false));
    assert_eq!(invoke("isError", vec![Value::Null]), Value::Bool(false));
}

// ── Error.prototype.toString (§20.5.3.4) ─────────────────────────────────────

#[test]
fn to_string_formats_name_colon_message() {
    use std::sync::Mutex;
    use vybe_runtime::value::Object;
    let e = invoke(
        "TypeError",
        vec![
            Value::Object(Arc::new(Mutex::new(Object::new()))),
            s("bad value"),
        ],
    );
    assert_eq!(
        invoke("toString", vec![e]),
        Value::String(Arc::from("TypeError: bad value"))
    );
}

#[test]
fn to_string_omits_message_when_empty() {
    use std::sync::Mutex;
    use vybe_runtime::value::Object;
    let e = invoke(
        "RangeError",
        vec![Value::Object(Arc::new(Mutex::new(Object::new()))), s("")],
    );
    assert_eq!(
        invoke("toString", vec![e]),
        Value::String(Arc::from("RangeError"))
    );
}

#[test]
fn to_string_of_plain_error_uses_error_name() {
    use std::sync::Mutex;
    use vybe_runtime::value::Object;
    let e = invoke(
        "Error",
        vec![
            Value::Object(Arc::new(Mutex::new(Object::new()))),
            s("oops"),
        ],
    );
    assert_eq!(
        invoke("toString", vec![e]),
        Value::String(Arc::from("Error: oops"))
    );
}
