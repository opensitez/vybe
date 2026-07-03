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
function AsyncFunction() { throw new TypeError("AsyncFunction constructor is not supported"); }
function GeneratorFunction() { throw new TypeError("GeneratorFunction constructor is not supported"); }
function AsyncGeneratorFunction() { throw new TypeError("AsyncGeneratorFunction constructor is not supported"); }

// These link fresh PER-VM objects (the hoisted declarations above), so they
// must run on every VM — NOT inside the prelude-done guard, whose flag rides
// the process-global shared prototypes and stays set across VMs in one
// process (test harness). All three statements are idempotent.
Object.setPrototypeOf(AsyncFunction.prototype, Function.prototype);
globalThis.AsyncFunction = AsyncFunction;
Object.setPrototypeOf(GeneratorFunction.prototype, Function.prototype);
globalThis.GeneratorFunction = GeneratorFunction;
Object.setPrototypeOf(AsyncGeneratorFunction.prototype, Function.prototype);
globalThis.AsyncGeneratorFunction = AsyncGeneratorFunction;

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
    let full_source = format!("{};\n{}", prelude, source);
    walker::parse(&full_source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}
