//! `vybe:object` — language-operator helpers without ECMA-262 equivalents.
//!
//! Most of `vybe:object.*` (keys, values, entries, assign, freeze, create,
//! seal, isFrozen, isSealed, is, getPrototypeOf, getOwnPropertyNames,
//! defineProperty, fromEntries, hasOwn, deleteProperty) was retired —
//! every one is now backed by the spec-correct `ecma:object.*` registered
//! under [`crate::ecma::object`]. Callers in language profiles and the
//! compiler emit `ecma:object` directly.
//!
//! Four entries remain because they implement language operators that
//! don't fit cleanly under ECMA-262 §19.1:
//!
//! - [`isset_all`] / [`is_empty`] — PHP `isset(...)` / `empty(...)`. Polymorphic
//!   "is this defined / is this falsy" primitives. PHP-specific value coercion
//!   (string `"0"` is empty, integer `0` is empty, etc.) doesn't map onto any
//!   ECMA predicate.
//! - [`hasProperty`] — `key in obj` operator. Argument order is `(key, obj)`,
//!   the OPPOSITE of `ecma:object.hasOwn(obj, key)`. Compilers emit calls
//!   in this order, so a direct alias swap would silently mis-bind. Kept
//!   as a thin language-operator helper until callers are migrated.
//! - [`instanceOf`] — `a instanceof B` operator. Walks the cross-language
//!   type registry (`__type` / `__types` / control-type metadata) to support
//!   instanceof across VB/JS/C# class hierarchies. Not an Object method
//!   per spec — the JS operator is its own AST node.

use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::ObjectKind;

pub fn register(vm: &mut VM) {
    // PHP `isset(a, b, c)` — true iff every arg is defined and non-null.
    // Defined as a polymorphic primitive so any language with a "are these
    // values all set?" test (Python `is not None` chains, JS `!= null`)
    // can reuse the same impl.
    vm.register_host_fn("vybe:object", "isset_all", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        for a in args {
            if matches!(a, Value::Null | Value::Undefined) {
                return Value::Bool(false);
            }
        }
        Value::Bool(!args.is_empty())
    }));

    // PHP `empty(v)` — true iff v is one of PHP's falsy values: null,
    // undefined, false, 0, 0.0, "", "0", empty array/map/set, or an
    // Object whose only own properties are internal `__`-prefixed metadata.
    vm.register_host_fn("vybe:object", "is_empty", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let Some(v) = args.first() else { return Value::Bool(true); };
        let falsy = match v {
            Value::Null | Value::Undefined => true,
            Value::Bool(b) => !b,
            Value::I32(n) => *n == 0,
            Value::I64(n) => *n == 0,
            Value::F64(n) => *n == 0.0,
            Value::String(s) => s.is_empty() || s.as_ref() == "0",
            Value::Object(obj) => {
                let o = obj.lock().unwrap();
                match &o.kind {
                    ObjectKind::Array(v) => v.is_empty(),
                    ObjectKind::Map(m) => m.is_empty(),
                    ObjectKind::Set(s) => s.is_empty(),
                    _ => o.properties.iter().all(|(k, _)| k.starts_with("__")),
                }
            }
            _ => false,
        };
        Value::Bool(falsy)
    }));

    // `key in obj` operator. NB: arg order is `(key, obj)`, not `(obj, key)`.
    // This is the opposite of `ecma:object.hasOwn` — kept here until the
    // compiler emits the canonical `(obj, key)` form directly to ecma:object.
    vm.register_host_fn("vybe:object", "hasProperty", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        if let Some(Value::Object(obj)) = args.get(1) {
            let o = obj.lock().unwrap();
            Value::Bool(o.properties.contains_key(&key))
        } else {
            Value::Bool(false)
        }
    }));

    // `a instanceof B` — cross-language type check. Walks `__type` /
    // `__types` / `__control_type` metadata stamped by the various class
    // emitters (VB designer codegen, JS class normalizer, dotnet ctors).
    vm.register_host_fn("vybe:object", "instanceOf", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let target_name = if let Some(Value::Object(ctor)) = args.get(1) {
            let ob = ctor.lock().unwrap();
            ob.properties.get("name").map(|v| format!("{}", v))
                .or_else(|| {
                    if let ObjectKind::Function(ref f) = ob.kind {
                        f.name.clone()
                    } else { None }
                })
                .unwrap_or_default()
        } else if let Some(Value::String(s)) = args.get(1) {
            s.to_string()
        } else {
            return Value::Bool(false);
        };
        if target_name.is_empty() { return Value::Bool(false); }

        if let Some(Value::Object(obj)) = args.first() {
            let o = obj.lock().unwrap();

            let obj_type_name = o.properties.get("__type")
                .map(|v| format!("{}", v))
                .or_else(|| o.properties.get("__control_type")
                    .map(|v| format!("{}", v)))
                .unwrap_or_default();

            if obj_type_name.eq_ignore_ascii_case(&target_name) {
                return Value::Bool(true);
            }

            // JS class inheritance chain stamped into __types by the class
            // normalizer.
            if let Some(Value::Object(types)) = o.properties.get("__types") {
                let t = types.lock().unwrap();
                if let ObjectKind::Array(ref elems) = t.kind {
                    if elems.iter().any(|e| format!("{}", e) == target_name) {
                        return Value::Bool(true);
                    }
                }
            }
        }
        Value::Bool(false)
    }));

}
