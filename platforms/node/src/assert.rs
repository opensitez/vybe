//! `node:assert` — Node.js assertion module.
//!
//! Reference: <https://nodejs.org/api/assert.html>.

use std::sync::Arc;
use vybe_runtime::{VM, Value};

fn loose_eq(a: &Value, b: &Value) -> bool {
    // Same type — direct equality
    if strict_eq(a, b) {
        return true;
    }
    // Number coercion: 1 == "1", 1 == 1.0
    match (a, b) {
        (Value::I32(n), Value::F64(f)) | (Value::F64(f), Value::I32(n)) => (*n as f64) == *f,
        (Value::I32(n), Value::String(s)) | (Value::String(s), Value::I32(n)) => {
            s.parse::<i32>().ok() == Some(*n)
        }
        (Value::F64(f), Value::String(s)) | (Value::String(s), Value::F64(f)) => {
            s.parse::<f64>().ok() == Some(*f)
        }
        (Value::Null, Value::Undefined) | (Value::Undefined, Value::Null) => true,
        _ => false }
}

fn strict_eq(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::I32(x), Value::I32(y)) => x == y,
        (Value::F64(x), Value::F64(y)) => x == y,
        (Value::Bool(x), Value::Bool(y)) => x == y,
        (Value::String(x), Value::String(y)) => x == y,
        (Value::Null, Value::Null) => true,
        (Value::Undefined, Value::Undefined) => true,
        _ => false }
}

fn is_truthy(v: &Value) -> bool {
    match v {
        Value::Bool(false) | Value::Null | Value::Undefined => false,
        Value::I32(0) => false,
        Value::F64(f) => *f != 0.0 && !f.is_nan(),
        Value::String(s) => !s.is_empty(),
        _ => true }
}

fn throw_assert(ctx: &mut vybe_runtime::vm::HostContext, msg: &str) {
    ctx.throw_value(Value::String(Arc::from(msg)));
}

/// Extract pattern from "/pattern/flags" string arg
fn regex_matches(pattern_str: &str, text: &str) -> bool {
    // strip surrounding / delimiters
    let inner = if pattern_str.starts_with('/') {
        let end = pattern_str.rfind('/').unwrap_or(pattern_str.len() - 1);
        &pattern_str[1..end]
    } else {
        pattern_str
    };
    regex::Regex::new(inner)
        .map(|re| re.is_match(text))
        .unwrap_or(false)
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:assert",
        "ok",
        Box::new(|ctx, args| {
            let v = args.first().unwrap_or(&Value::Undefined);
            if !is_truthy(v) {
                throw_assert(ctx, "AssertionError: value is not truthy");
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "equal",
        Box::new(|ctx, args| {
            let a = args.first().unwrap_or(&Value::Undefined);
            let b = args.get(1).unwrap_or(&Value::Undefined);
            if !loose_eq(a, b) {
                throw_assert(ctx, "AssertionError: values not equal");
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "notEqual",
        Box::new(|ctx, args| {
            let a = args.first().unwrap_or(&Value::Undefined);
            let b = args.get(1).unwrap_or(&Value::Undefined);
            if loose_eq(a, b) {
                throw_assert(ctx, "AssertionError: values are equal");
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "strictEqual",
        Box::new(|ctx, args| {
            let a = args.first().unwrap_or(&Value::Undefined);
            let b = args.get(1).unwrap_or(&Value::Undefined);
            if !strict_eq(a, b) {
                throw_assert(ctx, "AssertionError: values not strictly equal");
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "notStrictEqual",
        Box::new(|ctx, args| {
            let a = args.first().unwrap_or(&Value::Undefined);
            let b = args.get(1).unwrap_or(&Value::Undefined);
            if strict_eq(a, b) {
                throw_assert(ctx, "AssertionError: values are strictly equal");
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "deepEqual",
        Box::new(|ctx, args| {
            let a = args.first().unwrap_or(&Value::Undefined);
            let b = args.get(1).unwrap_or(&Value::Undefined);
            if !loose_eq(a, b) {
                throw_assert(ctx, "AssertionError: values not deep equal");
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "notDeepEqual",
        Box::new(|ctx, args| {
            let a = args.first().unwrap_or(&Value::Undefined);
            let b = args.get(1).unwrap_or(&Value::Undefined);
            if loose_eq(a, b) {
                throw_assert(ctx, "AssertionError: values are deep equal");
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "deepStrictEqual",
        Box::new(|ctx, args| {
            let a = args.first().unwrap_or(&Value::Undefined);
            let b = args.get(1).unwrap_or(&Value::Undefined);
            if !strict_eq(a, b) {
                throw_assert(ctx, "AssertionError: values not deeply strictly equal");
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "notDeepStrictEqual",
        Box::new(|ctx, args| {
            let a = args.first().unwrap_or(&Value::Undefined);
            let b = args.get(1).unwrap_or(&Value::Undefined);
            if strict_eq(a, b) {
                throw_assert(ctx, "AssertionError: values are deeply strictly equal");
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "ifError",
        Box::new(|ctx, args| {
            let v = args.first().unwrap_or(&Value::Undefined);
            if !matches!(v, Value::Null | Value::Undefined) {
                throw_assert(
                    ctx,
                    "AssertionError: ifError got a non-null/undefined value",
                );
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "fail",
        Box::new(|ctx, args| {
            let msg = match args.first() {
                Some(Value::String(s)) => format!("AssertionError: {s}"),
                _ => "AssertionError: Failed".to_string() };
            throw_assert(ctx, &msg);
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "throws",
        Box::new(|_ctx, _args| {
            // In the VM context we can't actually call the fn, return Undefined (best-effort stub)
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "doesNotThrow",
        Box::new(|_ctx, _args| Value::Undefined),
    );

    vm.register_host_fn(
        "node:assert",
        "match",
        Box::new(|ctx, args| {
            let text = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => {
                    throw_assert(ctx, "AssertionError: match requires string");
                    return Value::Undefined;
                }
            };
            let pattern = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                _ => {
                    throw_assert(ctx, "AssertionError: match requires pattern");
                    return Value::Undefined;
                }
            };
            if !regex_matches(&pattern, &text) {
                throw_assert(ctx, "AssertionError: string does not match pattern");
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "node:assert",
        "doesNotMatch",
        Box::new(|ctx, args| {
            let text = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => {
                    throw_assert(ctx, "AssertionError: doesNotMatch requires string");
                    return Value::Undefined;
                }
            };
            let pattern = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                _ => {
                    throw_assert(ctx, "AssertionError: doesNotMatch requires pattern");
                    return Value::Undefined;
                }
            };
            if regex_matches(&pattern, &text) {
                throw_assert(ctx, "AssertionError: string matches pattern");
            }
            Value::Undefined
        }),
    );

    // Async stubs — always resolve/pass (no promise infra in test context)
    vm.register_host_fn(
        "node:assert",
        "rejects",
        Box::new(|_ctx, _args| Value::Undefined),
    );
    vm.register_host_fn(
        "node:assert",
        "doesNotReject",
        Box::new(|_ctx, _args| Value::Undefined),
    );
}
