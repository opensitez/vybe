// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod emitter;
pub mod normalize_class;
pub mod protocol;
pub mod walker;

mod native_grammar {
    include!(concat!(env!("OUT_DIR"), "/native_parser.rs"));
}
pub(crate) use native_grammar::Rule;
pub(crate) struct JsParser;
impl JsParser {
    pub(crate) fn parse(rule: Rule, source: &str) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
        native_grammar::Parser.parse_pairs(rule, source)
    }
}

/// Parse JavaScript source into the common AST.
/// Parse ONLY the given source — no prelude prepend, no directive hoist —
/// so statement spans are in the CALLER's line/column coordinates. For
/// tooling that does span surgery on user text (eval's completion-value
/// extraction in vybex/dynamic.rs); execution still compiles via `parse`.
pub fn parse_source_only(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse_direct_recipes(source)
}

/// Explicit source-only PEG reference, preserving original caller coordinates.
pub fn parse_source_peg(source: &str) -> Result<vybe_ast::Module, String> { walker::parse(source) }

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
__vybe_wire_error_proto(SuppressedError, __vybe_error_proto_root, "SuppressedError");


if (!globalThis.__vybe_js_prelude_done) {
    globalThis.__vybe_js_prelude_done = true;






}
"#;
/// Parsed AST of [`JS_PRELUDE`], built once per process and borrowed per call.
/// The prelude is identical for every program, so re-parsing its ~284 lines
/// on every `parse` (7000× in the test suite) was pure waste — the dominant
/// per-test cost. A process-global cache is essential: the test harness
/// spawns a fresh thread per test, so a thread-local cache would be cold on
/// every test and re-parse anyway. Cloning the cached AST is far cheaper
/// than re-parsing. Entries are cloned directly into the final module body
/// to avoid allocating an intermediate prelude vector on each call.
fn prelude_body() -> &'static [vybe_ast::Statement] {
    // NOT on `vybe_compiler::primitives::prelude` like php/python/go: this
    // crate depends on `vybe_compiler` only as a DEV-dependency, and promoting
    // that to a real one would drag every JS build behind a compiler rebuild.
    // A single constant prelude needs one `OnceLock`, so the duplication is
    // one line rather than a mechanism.
    //
    // JS is also the one language whose prelude is UNCONDITIONAL — every
    // module gets all of it, because the function-kind intrinsics must be
    // hoisted before any user declaration stamps its metadata. Gating it is a
    // separate, measured change.
    static CACHE: std::sync::OnceLock<Vec<vybe_ast::Statement>> = std::sync::OnceLock::new();
    CACHE.get_or_init(|| {
        walker::parse_direct_bodies(JS_PRELUDE)
            .expect("JS prelude must parse")
            .body
    })
}

pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    // §11.2.1: the user's directive prologue must govern the module even
    // though the prelude runs first — surface a "use strict" directive at
    // the very front so the compiler's prologue scan (which stops at the
    // first non-directive statement) still sees it. Collect this fact during
    // source construction, before function hoisting changes statement order.
    // Parse ONLY the (small) user source, then splice in the cached prelude:
    // [ "use strict"?, prelude…, source… ]. Semantically identical to the old
    // textual prepend (the prelude defines globals the source references by
    // name at compile/runtime, never by AST identity), but skips re-parsing
    // the constant prelude every call.
    let (module, user_is_strict) = walker::parse_direct_recipes_with_prologue(source)?;
    Ok(prepend_prelude(module, user_is_strict))
}

/// Explicit PEG reference including the same strict-policy/prelude splice.
pub fn parse_peg(source: &str) -> Result<vybe_ast::Module, String> {
    let (module, strict) = walker::parse_with_source_prologue(source)?;
    Ok(prepend_prelude(module, strict))
}

/// Explicit arithmetic-only ownership for comparison with the broader frontend.
pub fn parse_direct_arithmetic(source: &str) -> Result<vybe_ast::Module, String> {
    let (module, strict) = walker::parse_direct_arithmetic_with_prologue(source)?;
    Ok(prepend_prelude(module, strict))
}

/// Explicit arithmetic and ordinary quoted-string ownership for comparison.
pub fn parse_direct_strings(source: &str) -> Result<vybe_ast::Module, String> {
    let (module, strict) = walker::parse_direct_strings_with_prologue(source)?;
    Ok(prepend_prelude(module, strict))
}

/// Retained owned body frontend reference; unsupported syntax retains captures.
pub fn parse_direct_bodies(source: &str) -> Result<vybe_ast::Module, String> {
    let (module, strict) = walker::parse_direct_bodies_with_prologue(source)?;
    Ok(prepend_prelude(module, strict))
}

/// Deferred expression recipes used by registered execution, with live-scope lowering.
pub fn parse_direct_recipes(source: &str) -> Result<vybe_ast::Module, String> {
    let (module, strict) = walker::parse_direct_recipes_with_prologue(source)?;
    Ok(prepend_prelude(module, strict))
}

fn prepend_prelude(mut module: vybe_ast::Module, user_is_strict: bool) -> vybe_ast::Module {
    let prelude = prelude_body();
    let mut body =
        Vec::with_capacity(prelude.len() + module.body.len() + 1 + usize::from(user_is_strict));
    if user_is_strict {
        body.push(vybe_ast::Statement::new(vybe_ast::StmtKind::Expr(
            vybe_ast::Expression::string("use strict"),
        )));
    }
    body.extend(prelude.iter().cloned());
    // Boundary marker so the step debugger can skip the runtime prelude and
    // land the first pause in the user's own code. The shared compiler
    // recognizes this sentinel string-expression, records the `<script>`
    // bytecode offset on `chunk.user_code_offset`, and emits nothing for it.
    // Generic mechanism — any frontend that injects a prelude can emit it.
    body.push(vybe_ast::Statement::new(vybe_ast::StmtKind::Expr(
        vybe_ast::Expression::string("__vybe_user_code_start__"),
    )));
    body.append(&mut module.body);
    module.body = body;
    module
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

fn compiler_optimizations(_selected: vybe_runtime::registry::ParseFn) -> vybe_runtime::registry::CompilerOptimizations {
    // Keep registration disabled until same-owner output and CPU gates pass.
    Default::default()
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    // Platforms this language needs. A language already links them (see
    // Cargo.toml), so registering them here means the compiler never has to —
    // and it works from ANY host, including language test binaries that never
    // construct vybex. See flexclassplan.md.
    vybe_platform_dotnet::register();

    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "js",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: None,
        expand_source: None,
    });
    vybe_runtime::registry::register_hooks(
        "js",
        vybe_runtime::registry::LanguageHooks {
            compiler_optimizations: Some(compiler_optimizations),
            proxy_get: Some(emitter::proxy_adapter::emit_proxy_get_dispatch),
            proxy_set: Some(emitter::proxy_adapter::emit_proxy_set_dispatch),
            proxy_set_bool: Some(emitter::proxy_adapter::emit_proxy_set_dispatch_bool),
            proxy_has: Some(emitter::proxy_adapter::emit_proxy_has_dispatch),
            proxy_create: Some(emitter::proxy_adapter::emit_proxy_create),
            proxy_delete_property: Some(
                emitter::proxy_adapter::emit_proxy_delete_property_dispatch,
            ),
            proxy_apply: Some(emitter::proxy_adapter::emit_proxy_apply_dispatch),
            proxy_own_keys: Some(emitter::proxy_adapter::emit_proxy_own_keys_dispatch),
            proxy_get_own_property_descriptor: Some(
                emitter::proxy_adapter::emit_proxy_get_own_property_descriptor_dispatch,
            ),
            proxy_define_property: Some(
                emitter::proxy_adapter::emit_proxy_define_property_dispatch,
            ),
            proxy_get_prototype_of: Some(
                emitter::proxy_adapter::emit_proxy_get_prototype_of_dispatch,
            ),
            proxy_set_prototype_of: Some(
                emitter::proxy_adapter::emit_proxy_set_prototype_of_dispatch,
            ),
            proxy_is_extensible: Some(emitter::proxy_adapter::emit_proxy_is_extensible_dispatch),
            proxy_prevent_extensions: Some(
                emitter::proxy_adapter::emit_proxy_prevent_extensions_dispatch,
            ),
            proxy_construct: Some(emitter::proxy_adapter::emit_proxy_construct_dispatch),
            parse_eval: Some(parse_source_only),
            ..Default::default()
        },
    );
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "js"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
    /// ⛔ `new Function(...)` STATE IS PER-TENANT AND WAS NEVER RECLAIMED.
    /// Each one boxes a whole `VM` into a process-global map in
    /// `vybe_compiler::dynamic` and leaks it: the map is thread-local, so
    /// `VM::reset_to` cannot reach it — which is precisely the case this hook
    /// exists for ("process-global state a plugin owns directly"). Without
    /// this, a warm worker carried every dynamic function every earlier
    /// program had built, and the id counter never restarted, so identical
    /// source lowered to different generated names depending on what had
    /// compiled before it.
    fn reset(&self) {
        vybe_compiler::dynamic::reset_dynamic_compilation_state();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);

#[cfg(test)]
mod public_prologue_tests {
    fn has_front_strict(source: &str) -> bool {
        let module = super::parse(source).unwrap_or_else(|error| panic!("{source:?}: {error}"));
        matches!(&module.body[0].kind,
            vybe_ast::StmtKind::Expr(expr)
            if matches!(&expr.kind, vybe_ast::ExprKind::Lit(vybe_ast::Literal::Str(value)) if value == "use strict"))
    }

    #[test]
    fn explicit_owned_body_frontend_preserves_public_prelude_scope_and_directives() {
        for source in [
            "'use asm'; 'use strict'; function f(){return document+2;}",
            "function f({x}=o){'use strict';return x+2;}",
            "class C extends B {m(x=2){'use strict';return x+3;}}",
            "function f(){function g(){return 2+3;} let x=4; return x+5;}",
            "const o = {get x(){return 2;},set x(v){v+3;}};",
            "function* f(){return 2+3;} const f2=()=>{return 2+3;};",
        ] {
            let peg = super::parse_peg(source).unwrap();
            assert_eq!(format!("{peg:?}"), format!("{:?}", super::parse_direct_arithmetic(source).unwrap()), "{source}");
            assert_eq!(format!("{peg:?}"), format!("{:?}", super::parse_direct_bodies(source).unwrap()), "{source}");
            assert_eq!(format!("{peg:?}"), format!("{:?}", super::parse_direct_recipes(source).unwrap()), "{source}");
        }
    }

    #[test]
    fn registered_hybrid_and_cached_prelude_match_explicit_peg_references() {
        assert_eq!(format!("{:?}",super::walker::parse(super::JS_PRELUDE)),
            format!("{:?}",super::walker::parse_direct_bodies(super::JS_PRELUDE)));
        for source in [
            "'use strict';function f(x=2){let y=3;if(x<y)return x+y;return y;}",
            "class C extends B {m(x=2){'use strict';return x>3;}}",
            "try {throw 2;}catch(e){const x=e+3;}finally {4+5;}",
        ] {
            assert_eq!(format!("{:?}",super::parse(source)),format!("{:?}",super::parse_peg(source)),"{source}");
            assert_eq!(format!("{:?}",super::parse_source_only(source)),format!("{:?}",super::parse_source_peg(source)),"{source}");
        }
    }

    #[test]
    fn expression_continuation_newlines_preserve_peg_ast() {
        for source in [
            "const x=new\nFoo();",
            "const x=typeof\ny;",
            "const x=!\ny;",
            "async function f(){return await\ng();}",
            "obj[\nkey\n];",
            "obj?.[\nkey\n];",
            "f(\n);",
            "function f(\n){return 1;}",
            "function f(\na,\nb\n){return a+b;}",
            "const f=(\na,\nb\n)=>a+b;",
            "var a =\n1;",
            "let a; a =\n2;",
            "let a=(x,\ny);",
            "obj.\nmethod();",
            "obj?.\nmethod();",
            "function f(a=\n1){return a;}",
            "let {a=\n1,b:c=\n2} = {};",
            "const x={a:1}; x.a =\n3;",
            "let a=1; a +=\n2;",
        ] {
            let peg = super::parse_source_peg(source).expect(source);
            let optimized = super::parse_source_only(source).expect(source);
            assert_eq!(format!("{peg:?}"), format!("{optimized:?}"), "{source}");
        }
    }

    #[test]
    fn restricted_throw_and_arrow_newlines_remain_errors() {
        for source in ["function f(){throw\n1;}", "x\n=>x;"] {
            let peg = super::parse_source_peg(source).expect_err(source);
            let optimized = super::parse_source_only(source).expect_err(source);
            assert_eq!(peg, optimized, "{source}");
        }
    }

    #[test]
    fn newline_and_comment_suffixes_keep_expected_ast_kinds() {
        use vybe_ast::{ExprKind, Expression, LambdaBody, Module, StmtKind};
        fn expression(module: &Module) -> &Expression {
            let StmtKind::Expr(expr) = &module.body[0].kind else { panic!("not expression") };
            expr
        }
        fn optional_branch(expr: &Expression) -> &Expression {
            let ExprKind::Call { callee, .. } = &expr.kind else { panic!("not lowered optional call") };
            let ExprKind::Lambda { body: LambdaBody::Expr(body), .. } = &callee.kind else { panic!("not optional lambda") };
            let ExprKind::Ternary { else_, .. } = &body.kind else { panic!("not optional ternary") };
            else_
        }
        let parse = |source: &str| {
            let peg = super::parse_source_peg(source).expect(source);
            let optimized = super::parse_source_only(source).expect(source);
            assert_eq!(format!("{peg:?}"), format!("{optimized:?}"), "{source}");
            optimized
        };
        for source in [
            "obj?.\n[key];", "obj?. /*c*/ [key];",
            "obj?. // c\r\n[key];", "obj?. /*\r\nc*/ [key];",
        ] {
            let module = parse(source);
            assert!(matches!(&optional_branch(expression(&module)).kind,
                ExprKind::Index { index, .. } if matches!(&index.kind, ExprKind::Ident(s) if s == "key")), "{source}");
        }
        for source in ["obj?.\n();", "obj?. /*c*/ ();", "obj?.\n(x);", "obj?. // c\r\n();"] {
            let module = parse(source);
            assert!(matches!(&optional_branch(expression(&module)).kind, ExprKind::Call { .. }), "{source}");
        }
        for source in ["obj\n/*c*/ .field;", "obj\n// c\n.field;", "obj\r\n// c\r\n.field;"] {
            let module = parse(source);
            assert!(matches!(&expression(&module).kind, ExprKind::Member { field, .. } if field == "field"), "{source}");
        }
        let module = parse("new Foo\n/*c*/ .Bar();");
        assert!(matches!(&expression(&module).kind, ExprKind::New { class, .. }
            if matches!(&class.kind, ExprKind::Member { field, .. } if field == "Bar")));
    }

    #[test]
    fn full_source_directive_prologue_survives_prelude() {
        for source in [
            "'use strict'; function f() { return this; }",
            "'use asm'; 'use strict'; function f() { return this; }",
            "'other'; /* between */ \"use strict\"; var x = 1;",
            "// leading\n'other'\n'use strict'\nvar x = 1;",
            "#!/usr/bin/env node\n'use asm'; 'use strict';",
        ] {
            assert!(has_front_strict(source), "{source:?}");
        }
    }

    #[test]
    fn only_raw_unparenthesized_initial_string_statements_are_directives() {
        for source in [
            "'use strict' + 'suffix';",
            "('use strict');",
            "`use strict`;",
            "'use\\x20strict';",
            "; 'use strict';",
            "'other';; 'use strict';",
            "var x = 1; 'use strict';",
            "function f() {} 'use strict';",
            "function f() { 'use strict'; }",
        ] {
            assert!(!has_front_strict(source), "{source:?}");
        }
    }
}
