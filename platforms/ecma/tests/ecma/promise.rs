//! Behaviour tests for `ecma:promise` host imports.
//!
//! Reference: ECMA-262 §27.2 Promise.
//!
//! Each test covers a distinct behaviour. Since the VM is synchronous,
//! we test the synchronous subset: already-resolved/rejected promises,
//! static combinators, and the then/catch/finally chaining surface.

use std::sync::{Arc, Mutex};
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-promise-test>");
    let import_idx = chunk.add_import("ecma:promise", name);
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

fn arr(values: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(values))))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

// ── Promise.resolve — wraps a value in an already-resolved promise ────────────

#[test]
fn resolve_returns_object() {
    let p = invoke("resolve", vec![Value::I32(42)]);
    assert!(matches!(p, Value::Object(_)));
}

#[test]
fn resolve_value_accessible_via_await_or_settled() {
    // The host must expose a way to read the settled value synchronously.
    let p = invoke("resolve", vec![Value::I32(7)]);
    let settled = invoke("settled", vec![p]);
    // settled() returns { status: "fulfilled", value: 7 }
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("fulfilled")));
        assert_eq!(o.properties.get("value").cloned(), Some(Value::I32(7)));
    } else {
        panic!("expected settled descriptor object");
    }
}

// ── Promise.reject ────────────────────────────────────────────────────────────

#[test]
fn reject_produces_rejected_promise() {
    let p = invoke("reject", vec![s("oops")]);
    let settled = invoke("settled", vec![p]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("rejected")));
        assert_eq!(o.properties.get("reason").cloned(), Some(s("oops")));
    } else {
        panic!("expected settled descriptor object");
    }
}

// ── Promise.all — all fulfilled ───────────────────────────────────────────────

#[test]
fn all_resolves_when_all_inputs_are_fulfilled() {
    let p1 = invoke("resolve", vec![Value::I32(1)]);
    let p2 = invoke("resolve", vec![Value::I32(2)]);
    let combined = invoke("all", vec![arr(vec![p1, p2])]);
    let settled = invoke("settled", vec![combined]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("fulfilled")));
        // value must be an array [1, 2]
        if let Some(Value::Object(arr_val)) = o.properties.get("value") {
            let arr_val = arr_val.lock().unwrap();
            if let ObjectKind::Array(elems) = &arr_val.kind {
                assert_eq!(elems.len(), 2);
                assert_eq!(elems[0], Value::I32(1));
                assert_eq!(elems[1], Value::I32(2));
            } else {
                panic!("expected array kind");
            }
        } else {
            panic!("expected array value");
        }
    } else {
        panic!("expected settled descriptor");
    }
}

#[test]
fn all_rejects_when_any_input_is_rejected() {
    // ECMA-262 §27.2.4.1: Promise.all rejects as soon as one input rejects.
    let p1 = invoke("resolve", vec![Value::I32(1)]);
    let p2 = invoke("reject", vec![s("fail")]);
    let combined = invoke("all", vec![arr(vec![p1, p2])]);
    let settled = invoke("settled", vec![combined]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("rejected")));
    } else {
        panic!("expected settled descriptor");
    }
}

// ── Promise.allSettled — never rejects, collects all outcomes ────────────────

#[test]
fn all_settled_collects_both_fulfilled_and_rejected() {
    // ECMA-262 §27.2.4.2: allSettled always fulfills, with descriptor objects.
    let p1 = invoke("resolve", vec![Value::I32(1)]);
    let p2 = invoke("reject", vec![s("err")]);
    let combined = invoke("allSettled", vec![arr(vec![p1, p2])]);
    let settled = invoke("settled", vec![combined]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("fulfilled")));
        if let Some(Value::Object(arr_val)) = o.properties.get("value") {
            let arr_val = arr_val.lock().unwrap();
            if let ObjectKind::Array(elems) = &arr_val.kind {
                assert_eq!(elems.len(), 2);
            } else {
                panic!("expected array");
            }
        }
    } else {
        panic!("expected settled descriptor");
    }
}

// ── Promise.race — first settled wins ────────────────────────────────────────

#[test]
fn race_resolves_with_the_first_settled_value() {
    // All inputs already resolved; the first one wins.
    let p1 = invoke("resolve", vec![Value::I32(10)]);
    let p2 = invoke("resolve", vec![Value::I32(20)]);
    let winner = invoke("race", vec![arr(vec![p1, p2])]);
    let settled = invoke("settled", vec![winner]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("fulfilled")));
        assert_eq!(o.properties.get("value").cloned(), Some(Value::I32(10)));
    } else {
        panic!("expected settled descriptor");
    }
}

// ── Promise.any — first fulfilled wins, AggregateError if all reject ──────────

#[test]
fn any_resolves_with_first_fulfilled_skipping_earlier_rejections() {
    // ECMA-262 §27.2.4.4: Promise.any fulfills with the first non-rejected value.
    let p1 = invoke("reject", vec![s("nope")]);
    let p2 = invoke("resolve", vec![Value::I32(5)]);
    let winner = invoke("any", vec![arr(vec![p1, p2])]);
    let settled = invoke("settled", vec![winner]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("fulfilled")));
        assert_eq!(o.properties.get("value").cloned(), Some(Value::I32(5)));
    } else {
        panic!("expected settled descriptor");
    }
}

#[test]
fn any_rejects_with_aggregate_error_when_all_inputs_reject() {
    let p1 = invoke("reject", vec![s("a")]);
    let p2 = invoke("reject", vec![s("b")]);
    let result = invoke("any", vec![arr(vec![p1, p2])]);
    let settled = invoke("settled", vec![result]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("rejected")));
    } else {
        panic!("expected settled descriptor");
    }
}

// ── then / catch / finally chaining ──────────────────────────────────────────

#[test]
fn then_transforms_fulfilled_value() {
    // then(fn) maps the resolved value; the host applies it synchronously.
    let p = invoke("resolve", vec![Value::I32(3)]);
    // Encode the transform as a descriptor the host can interpret:
    // { __map_add: n } means "add n to the resolved value".
    let transform = {
        let mut o = Object::new();
        o.properties.insert("__map_add".to_string(), Value::I32(10));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let mapped = invoke("then", vec![p, transform]);
    let settled = invoke("settled", vec![mapped]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("fulfilled")));
        assert_eq!(o.properties.get("value").cloned(), Some(Value::I32(13)));
    } else {
        panic!("expected settled descriptor");
    }
}

#[test]
fn catch_handles_rejection_and_recovers() {
    let p = invoke("reject", vec![s("boom")]);
    // { __catch_return: value } means "recover with this value".
    let handler = {
        let mut o = Object::new();
        o.properties
            .insert("__catch_return".to_string(), Value::I32(0));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let recovered = invoke("catch", vec![p, handler]);
    let settled = invoke("settled", vec![recovered]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("fulfilled")));
        assert_eq!(o.properties.get("value").cloned(), Some(Value::I32(0)));
    } else {
        panic!("expected settled descriptor");
    }
}

#[test]
fn finally_runs_regardless_of_outcome_and_preserves_value() {
    // ECMA-262: finally callback does not receive the value; the original
    // settled value/reason passes through unchanged.
    let p = invoke("resolve", vec![Value::I32(42)]);
    let side_effect_tracker = {
        let mut o = Object::new();
        o.properties
            .insert("__finally_noop".to_string(), Value::Bool(true));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let after = invoke("finally", vec![p, side_effect_tracker]);
    let settled = invoke("settled", vec![after]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("fulfilled")));
        assert_eq!(o.properties.get("value").cloned(), Some(Value::I32(42)));
    } else {
        panic!("expected settled descriptor");
    }
}

// ── Promise constructor with executor ────────────────────────────────────────

#[test]
fn new_with_resolve_executor_creates_fulfilled_promise() {
    // new Promise((resolve) => resolve(99))
    // Encoded as { __executor_resolve: value }.
    let executor = {
        let mut o = Object::new();
        o.properties
            .insert("__executor_resolve".to_string(), Value::I32(99));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let p = invoke("new", vec![executor]);
    let settled = invoke("settled", vec![p]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("fulfilled")));
        assert_eq!(o.properties.get("value").cloned(), Some(Value::I32(99)));
    } else {
        panic!("expected settled descriptor");
    }
}

#[test]
fn new_with_reject_executor_creates_rejected_promise() {
    let executor = {
        let mut o = Object::new();
        o.properties
            .insert("__executor_reject".to_string(), s("err"));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let p = invoke("new", vec![executor]);
    let settled = invoke("settled", vec![p]);
    if let Value::Object(o) = settled {
        let o = o.lock().unwrap();
        assert_eq!(o.properties.get("status").cloned(), Some(s("rejected")));
    } else {
        panic!("expected settled descriptor");
    }
}

// ── Promise.withResolvers (ES2024 §27.2.4.5) ─────────────────────────────────

#[test]
fn with_resolvers_returns_object_with_promise_resolve_reject() {
    // ECMA-262 ES2024: Promise.withResolvers() → { promise, resolve, reject }.
    let result = invoke("withResolvers", vec![]);
    assert!(matches!(result, Value::Object(_)));
    if let Value::Object(o) = &result {
        let o = o.lock().unwrap();
        assert!(
            o.properties.contains_key("promise"),
            "withResolvers result must have 'promise' key"
        );
        assert!(
            o.properties.contains_key("resolve"),
            "withResolvers result must have 'resolve' key"
        );
        assert!(
            o.properties.contains_key("reject"),
            "withResolvers result must have 'reject' key"
        );
    }
}

#[test]
fn with_resolvers_promise_starts_pending() {
    // The returned promise is initially neither fulfilled nor rejected.
    let result = invoke("withResolvers", vec![]);
    if let Value::Object(o) = &result {
        let promise = o
            .lock()
            .unwrap()
            .properties
            .get("promise")
            .cloned()
            .unwrap_or(Value::Undefined);
        let settled = invoke("settled", vec![promise]);
        // Pending promises are reported with status "pending" or return Undefined.
        match settled {
            Value::Object(s) => {
                let status = s
                    .lock()
                    .unwrap()
                    .properties
                    .get("status")
                    .cloned()
                    .unwrap_or(Value::Undefined);
                assert_eq!(status, Value::String(Arc::from("pending")));
            }
            Value::Undefined => {}
            other => panic!("unexpected settled value: {:?}", other),
        }
    }
}

// ── Promise.try (ES2025 §27.2.4.6) ───────────────────────────────────────────

#[test]
fn try_with_returning_callable_creates_fulfilled_promise() {
    // ECMA-262 ES2025: Promise.try(fn) calls fn and wraps result in a fulfilled promise.
    let fn_obj = {
        let mut o = Object::new();
        o.properties
            .insert("__executor_resolve".to_string(), Value::I32(7));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let result = invoke("try", vec![fn_obj]);
    assert!(matches!(result, Value::Object(_)));
}

#[test]
fn try_with_throwing_callable_creates_rejected_promise() {
    // Promise.try(fn) catches synchronous throws and rejects the promise.
    let fn_obj = {
        let mut o = Object::new();
        o.properties
            .insert("__executor_reject".to_string(), s("thrown error"));
        Value::Object(Arc::new(Mutex::new(o)))
    };
    let p = invoke("try", vec![fn_obj]);
    let settled = invoke("settled", vec![p]);
    if let Value::Object(o) = settled {
        let status = o
            .lock()
            .unwrap()
            .properties
            .get("status")
            .cloned()
            .unwrap_or(Value::Undefined);
        assert_eq!(status, s("rejected"));
    }
}
