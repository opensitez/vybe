pub mod emitter;
pub mod normalize_class;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct JsParser;

/// Parse JavaScript source into the common AST.
/// Parse ONLY the given source — no prelude prepend, no directive hoist —
/// so statement spans are in the CALLER's line/column coordinates. For
/// tooling that does span surgery on user text (eval's completion-value
/// extraction in vybex/dynamic.rs); execution still compiles via `parse`.
pub fn parse_source_only(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// The JS runtime prelude — intrinsics, Error prototypes, Map/Set wrappers,
/// `Symbol.hasInstance`, … — textually the first thing every module runs.
/// Constant across all programs, so it is parsed ONCE (see `prelude_body`)
/// rather than re-parsed on every `parse` call.
const JS_PRELUDE: &str = r#"
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
/// Parsed AST of [`JS_PRELUDE`], built once per PROCESS and cloned per call.
/// The prelude is identical for every program, so re-parsing its ~284 lines
/// on every `parse` (7000× in the test suite) was pure waste — the dominant
/// per-test cost. A process-global cache is essential: the test harness
/// spawns a fresh thread per test, so a thread-local cache would be cold on
/// every test and re-parse anyway. Cloning the cached AST is far cheaper
/// than re-parsing.
fn prelude_body() -> Vec<vybe_ast::Statement> {
    static CACHE: std::sync::OnceLock<Vec<vybe_ast::Statement>> = std::sync::OnceLock::new();
    CACHE
        .get_or_init(|| {
            walker::parse(JS_PRELUDE)
                .expect("JS prelude must parse")
                .body
        })
        .clone()
}

pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    // §11.2.1: the user's directive prologue must govern the module even
    // though the prelude runs first — surface a "use strict" directive at
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

    // Parse ONLY the (small) user source, then splice in the cached prelude:
    // [ "use strict"?, prelude…, source… ]. Semantically identical to the old
    // textual prepend (the prelude defines globals the source references by
    // name at compile/runtime, never by AST identity), but skips re-parsing
    // the constant prelude every call.
    let mut module = walker::parse(source)?;
    let prelude = prelude_body();
    let mut body = Vec::with_capacity(prelude.len() + module.body.len() + 1);
    if user_is_strict {
        body.push(vybe_ast::Statement::new(vybe_ast::StmtKind::Expr(
            vybe_ast::Expression::string("use strict"),
        )));
    }
    body.extend(prelude);
    body.append(&mut module.body);
    module.body = body;
    Ok(module)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    vybe_plugin::registry::register_language(vybe_plugin::registry::LanguagePlugin {
        name: "js",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: None,
    });
    vybe_plugin::registry::register_hooks("js", vybe_plugin::registry::LanguageHooks {
        proxy_get: Some(emitter::proxy_adapter::emit_proxy_get_dispatch),
        proxy_set: Some(emitter::proxy_adapter::emit_proxy_set_dispatch),
        proxy_set_bool: Some(emitter::proxy_adapter::emit_proxy_set_dispatch_bool),
        proxy_has: Some(emitter::proxy_adapter::emit_proxy_has_dispatch),
        proxy_create: Some(emitter::proxy_adapter::emit_proxy_create),
        parse_eval: Some(parse_source_only),
        ..Default::default()
    });
}

/// This crate as a [`vybe_plugin::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_plugin::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "js"
    }
    fn init(&self, _fw: &mut vybe_plugin::Framework<'_>) {
        register();
    }
}
