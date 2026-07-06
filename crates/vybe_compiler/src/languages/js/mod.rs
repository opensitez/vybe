pub mod emitter;
pub mod normalize_class;
mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "languages/js/grammar.pest"]
pub(crate) struct JsParser;

/// Parse JavaScript source into the common AST.
pub fn parse(source: &str) -> Result<crate::ast::Module, String> {
    let prelude = r#"
// Function-kind intrinsics (ECMA-262 %AsyncFunction% §27.7.1,
// %GeneratorFunction% §27.3.1, %AsyncGeneratorFunction% §27.4.1). These
// declarations MUST be top-level (unguarded): user function declarations
// hoist above the guarded prelude body, and their metadata stamps read
// these globals — a guarded declaration would not exist yet at that point.
// The declaration-created `.prototype` objects are mutated in place
// (setPrototypeOf inside the guard), never replaced, so hoisted stamps
// keep identity. Each prototype inherits from %Function.prototype%, so
// `asyncFn instanceof Function` stays true through the chain.
// §27.3.2 / §27.7.3 / §27.4.2: each intrinsic constructor's `length` is 1
// (the trailing body argument) — hence the single parameter.
function AsyncFunction(body) { throw new TypeError("AsyncFunction constructor is not supported"); }
function GeneratorFunction(body) { throw new TypeError("GeneratorFunction constructor is not supported"); }
function AsyncGeneratorFunction(body) { throw new TypeError("AsyncGeneratorFunction constructor is not supported"); }

// These link fresh PER-VM objects (the hoisted declarations above), so they
// must run on every VM — NOT inside the prelude-done guard, whose flag rides
// the process-global shared prototypes and stays set across VMs in one
// process (test harness). All three statements are idempotent.
Object.setPrototypeOf(AsyncFunction.prototype, Function.prototype);
Object.setPrototypeOf(GeneratorFunction.prototype, Function.prototype);
Object.setPrototypeOf(AsyncGeneratorFunction.prototype, Function.prototype);
// §19: well-known intrinsics on the global object are non-enumerable.
Object.defineProperty(globalThis, "AsyncFunction", { value: AsyncFunction, writable: true, enumerable: false, configurable: true });
Object.defineProperty(globalThis, "GeneratorFunction", { value: GeneratorFunction, writable: true, enumerable: false, configurable: true });
Object.defineProperty(globalThis, "AsyncGeneratorFunction", { value: AsyncGeneratorFunction, writable: true, enumerable: false, configurable: true });

// §20.2.3.6 Function.prototype[Symbol.hasInstance] — OrdinaryHasInstance:
// walk V's prototype chain looking for this.prototype. Lives on
// %Function.prototype% so EVERY function inherits the SAME fn
// (F[Symbol.hasInstance] === Function[Symbol.hasInstance]); also set on
// Function directly since the global constructor object may lack the
// proto link. Must stay idempotent (configurable: true): this unguarded
// prelude runs once per VM and several VMs share one process (test
// harness) — a non-configurable redefine would throw on the second VM.
var __vybe_ordinary_has_instance = function(v) {
    if (v === null || v === undefined) return false;
    let p = Object.getPrototypeOf(Object(v));
    const target = this.prototype;
    while (p) { if (p === target) return true; p = Object.getPrototypeOf(p); }
    return false;
};
Object.defineProperty(Function.prototype, Symbol.hasInstance, { value: __vybe_ordinary_has_instance, writable: false, enumerable: false, configurable: true });
Object.defineProperty(Function, Symbol.hasInstance, { value: __vybe_ordinary_has_instance, writable: false, enumerable: false, configurable: true });

// §20.5.3 %Error.prototype% + §20.5.6.3 NativeError prototypes — declared
// here in JS so the whole prototype chain is pure WASM state. Bare
// `Error`/`TypeError`/… resolve to the canonical `__ctor_<Name>` anchors;
// the `ecma:error` host constructors link every error they mint to these
// same per-VM objects, so host-minted and compiled errors share one chain.
// Runs unguarded: the ctor anchors are per-VM.
// NOTE: the toString value below MUST stay an ANONYMOUS function — a
// named function expression whose name collides with a universal method
// name ("toString") poisons compile-time bookkeeping and breaks later
// object-literal method emission (generator methods vanish).
var __vybe_wire_error_proto = function(C, parentProto, name) {
    const proto = Object.create(parentProto);
    Object.defineProperty(proto, "constructor", { value: C, writable: true, enumerable: false, configurable: true });
    Object.defineProperty(proto, "name", { value: name, writable: true, enumerable: false, configurable: true });
    Object.defineProperty(proto, "message", { value: "", writable: true, enumerable: false, configurable: true });
    C.prototype = proto;
    return proto;
};
var __vybe_error_proto_root = __vybe_wire_error_proto(Error, Object.prototype, "Error");
// §20.5.3.4 Error.prototype.toString
Object.defineProperty(__vybe_error_proto_root, "toString", { value: function () {
    let n = this.name;
    n = n === undefined ? "Error" : "" + n;
    let m = this.message;
    m = m === undefined ? "" : "" + m;
    if (n === "") return m;
    if (m === "") return n;
    return n + ": " + m;
}, writable: true, enumerable: false, configurable: true });
__vybe_wire_error_proto(TypeError, __vybe_error_proto_root, "TypeError");
__vybe_wire_error_proto(RangeError, __vybe_error_proto_root, "RangeError");
__vybe_wire_error_proto(ReferenceError, __vybe_error_proto_root, "ReferenceError");
__vybe_wire_error_proto(SyntaxError, __vybe_error_proto_root, "SyntaxError");
__vybe_wire_error_proto(URIError, __vybe_error_proto_root, "URIError");
__vybe_wire_error_proto(EvalError, __vybe_error_proto_root, "EvalError");
__vybe_wire_error_proto(AggregateError, __vybe_error_proto_root, "AggregateError");


if (!globalThis.__vybe_js_prelude_done) {
    globalThis.__vybe_js_prelude_done = true;

    const _originalMap = globalThis.Map;
    const _originalSet = globalThis.Set;
    const _originalBoolean = globalThis.Boolean;
    const _originalNumber = globalThis.Number;

    const _originalIsPrototypeOf = Object.prototype.isPrototypeOf;
    Object.prototype.isPrototypeOf = function(v) {
        if (v === null || v === undefined) return false;
        return _originalIsPrototypeOf.call(this, Object(v));
    };

    if (_originalMap) {
        const MapConstructor = function(...args) {
            if (!new.target) throw new TypeError("Constructor Map requires 'new'");
            const instance = _originalMap.new();
            Object.setPrototypeOf(instance, MapConstructor.prototype);
            if (args[0] !== undefined && args[0] !== null) {
                for (const item of args[0]) {
                    instance.set(item[0], item[1]);
                }
            }
            return instance;
        };
        for (const k of Object.getOwnPropertyNames(_originalMap)) {
            if (k !== 'new') {
                Object.defineProperty(MapConstructor, k, {
                    value: _originalMap[k],
                    writable: true,
                    enumerable: false,
                    configurable: true
                });
            }
        }
        MapConstructor.prototype = Object.create(Object.prototype);
        Object.defineProperty(MapConstructor.prototype, 'constructor', {
            value: MapConstructor,
            writable: true,
            enumerable: false,
            configurable: true
        });
        MapConstructor.prototype[Symbol.toStringTag] = "Map";
        Object.defineProperty(MapConstructor.prototype, Symbol.toStringTag, { enumerable: false });
        for (const k of ['get', 'set', 'has', 'delete', 'clear', 'keys', 'values', 'entries', 'forEach']) {
            Object.defineProperty(MapConstructor.prototype, k, {
                value: _originalMap[k],
                writable: true,
                enumerable: false,
                configurable: true
            });
        }
        Object.defineProperty(MapConstructor.prototype, 'size', {
            get: function() { return _originalMap.size(this); },
            enumerable: false,
            configurable: true
        });
        globalThis.Map = MapConstructor;
    }

    if (_originalSet) {
        const SetConstructor = function(...args) {
            if (!new.target) throw new TypeError("Constructor Set requires 'new'");
            const instance = _originalSet.new();
            Object.setPrototypeOf(instance, SetConstructor.prototype);
            if (args[0] !== undefined && args[0] !== null) {
                for (const item of args[0]) {
                    instance.add(item);
                }
            }
            return instance;
        };
        for (const k of Object.getOwnPropertyNames(_originalSet)) {
            if (k !== 'new') {
                Object.defineProperty(SetConstructor, k, {
                    value: _originalSet[k],
                    writable: true,
                    enumerable: false,
                    configurable: true
                });
            }
        }
        SetConstructor.prototype = Object.create(Object.prototype);
        Object.defineProperty(SetConstructor.prototype, 'constructor', {
            value: SetConstructor,
            writable: true,
            enumerable: false,
            configurable: true
        });
        SetConstructor.prototype[Symbol.toStringTag] = "Set";
        Object.defineProperty(SetConstructor.prototype, Symbol.toStringTag, { enumerable: false });
        for (const k of ['add', 'has', 'delete', 'clear', 'keys', 'values', 'entries', 'forEach']) {
            Object.defineProperty(SetConstructor.prototype, k, {
                value: _originalSet[k],
                writable: true,
                enumerable: false,
                configurable: true
            });
        }
        Object.defineProperty(SetConstructor.prototype, 'size', {
            get: function() { return _originalSet.size(this); },
            enumerable: false,
            configurable: true
        });
        globalThis.Set = SetConstructor;
    }

    if (_originalBoolean) {
        Boolean.prototype[Symbol.toStringTag] = "Boolean";
        Object.defineProperty(Boolean.prototype, Symbol.toStringTag, { enumerable: false });
    }

    if (_originalNumber) {
        Number.prototype[Symbol.toStringTag] = "Number";
        Object.defineProperty(Number.prototype, Symbol.toStringTag, { enumerable: false });

        const _originalToFixed = Number.prototype.toFixed;
        Number.prototype.toFixed = function(fractionDigits) {
            const val = Number(this);
            if (!isFinite(val)) {
                if (isNaN(val)) return "NaN";
                return val > 0 ? "Infinity" : "-Infinity";
            }
            if (fractionDigits !== undefined) {
                const f = Number(fractionDigits);
                if (f < 0 || f > 100 || isNaN(f)) {
                    throw new RangeError("toFixed() digits out of range");
                }
            }
            if (Object.is(val, -0) || Object.is(val, 0)) {
                const temp = _originalToFixed.call(this, fractionDigits);
                if (temp.startsWith("-")) return temp.substring(1);
                return temp;
            }
            const res = _originalToFixed.call(this, fractionDigits);
            if (res === "inf") return "Infinity";
            if (res === "-inf") return "-Infinity";
            if (res === "nan") return "NaN";
            return res;
        };

        const _originalToExponential = Number.prototype.toExponential;
        Number.prototype.toExponential = function(fractionDigits) {
            const val = Number(this);
            if (!isFinite(val)) {
                if (isNaN(val)) return "NaN";
                return val > 0 ? "Infinity" : "-Infinity";
            }
            if (fractionDigits !== undefined) {
                const f = Number(fractionDigits);
                if (f < 0 || f > 100 || isNaN(f)) {
                    throw new RangeError("toExponential() digits out of range");
                }
            }
            const res = _originalToExponential.call(this, fractionDigits);
            if (res === "inf") return "Infinity";
            if (res === "-inf") return "-Infinity";
            if (res === "nan") return "NaN";
            return res;
        };

        const _originalToPrecision = Number.prototype.toPrecision;
        Number.prototype.toPrecision = function(precision) {
            const val = Number(this);
            if (!isFinite(val)) {
                if (isNaN(val)) return "NaN";
                return val > 0 ? "Infinity" : "-Infinity";
            }
            if (precision !== undefined) {
                const p = Number(precision);
                if (p < 1 || p > 100 || isNaN(p)) {
                    throw new RangeError("toPrecision() precision out of range");
                }
            }
            const res = _originalToPrecision.call(this, precision);
            if (res === "inf") return "Infinity";
            if (res === "-inf") return "-Infinity";
            if (res === "nan") return "NaN";
            return res;
        };

        const _originalNumberToString = Number.prototype.toString;
        Number.prototype.toString = function(radix) {
            const val = Number(this);
            if (radix !== undefined) {
                const r = Number(radix);
                if (r < 2 || r > 36 || isNaN(r)) {
                    throw new RangeError("toString() radix must be between 2 and 36");
                }
            }
            const res = _originalNumberToString.call(this, radix);
            if (res === "inf") return "Infinity";
            if (res === "-inf") return "-Infinity";
            if (res === "nan") return "NaN";
            return res;
        };
    }
}
"#;
    // §11.2.1: the user's directive prologue must govern the module even
    // though the prelude is textually prepended — hoist "use strict" to
    // the very front so the compiler's prologue scan (which stops at the
    // first non-directive statement) still sees it.
    let user_is_strict = {
        let mut t = source;
        loop {
            let trimmed = t.trim_start();
            if let Some(rest) = trimmed.strip_prefix("//") {
                t = rest.split_once('\n').map(|x| x.1).unwrap_or("");
            } else if let Some(rest) = trimmed.strip_prefix("/*") {
                t = rest.split_once("*/").map(|x| x.1).unwrap_or("");
            } else {
                t = trimmed;
                break;
            }
        }
        t.starts_with("\"use strict\"") || t.starts_with("'use strict'")
    };
    let full_source = if user_is_strict {
        format!("\"use strict\";\n{};\n{}", prelude, source)
    } else {
        format!("{};\n{}", prelude, source)
    };
    walker::parse(&full_source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}
