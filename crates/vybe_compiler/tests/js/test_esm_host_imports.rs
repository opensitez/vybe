//! ESM host-module imports — the spec-compliant forms.
//!
//! `import { X } from "wasi:foo"` should bind X at compile time so:
//!   * `X(args)` compiles to `CALL_IMPORT` (no runtime lookup)
//!   * `const f = X; f(args)` reads X as a value (runtime global)
//!
//! `import { X as Y } from "wasi:foo"` — local binding Y.
//! `import * as ns from "wasi:foo"` — namespace object with all host fns.
//! `import "wasi:foo"` — side-effect only (no-op for host modules).
//! `import X from "wasi:foo"` — default, no meaning for host modules (no-op).

use super::helpers::run_js_with_imports;

#[test]
fn named_import_direct_call() {
    let out = run_js_with_imports(r#"
import { log } from "wasi:cli";
log("direct");
"#);
    assert_eq!(out, vec!["direct"]);
}

#[test]
fn named_import_read_as_value() {
    // `const f = log; f(...)` — exercises the GLOBAL_GET path.
    let out = run_js_with_imports(r#"
import { log } from "wasi:cli";
const f = log;
f("indirect");
"#);
    assert_eq!(out, vec!["indirect"]);
}

#[test]
fn named_import_aliased() {
    let out = run_js_with_imports(r#"
import { log as myLog } from "wasi:cli";
myLog("aliased");
"#);
    assert_eq!(out, vec!["aliased"]);
}

#[test]
fn wildcard_namespace_import() {
    // `import * as cli from "wasi:cli"` synthesizes a namespace object
    // with `log` (and every other fn in wasi:cli) as properties.
    let out = run_js_with_imports(r#"
import * as cli from "wasi:cli";
cli.log("namespaced");
"#);
    assert_eq!(out, vec!["namespaced"]);
}

#[test]
fn side_effect_import_is_noop() {
    // Side-effect import of a host module must not error — host modules
    // aren't code that runs, they're a bag of functions.
    let out = run_js_with_imports(r#"
import "wasi:cli";
console.log("after");
"#);
    assert_eq!(out, vec!["after"]);
}

// ── Phase 5: Module Namespace Exotic Object (ECMA-262 §10.4.6) ─────

#[test]
fn wildcard_namespace_typeof_is_object() {
    // Per ECMA-262 §10.4.6, `typeof ns === "object"` even though the
    // object is exotic (frozen, null-prototype).
    let out = run_js_with_imports(r#"
import * as cli from "wasi:cli";
console.log(typeof cli);
"#);
    assert_eq!(out, vec!["object"]);
}

#[test]
fn wildcard_namespace_tostring_tag() {
    // `Object.prototype.toString.call(ns)` === `"[object Module]"`
    // courtesy of the `Symbol.toStringTag` own property = `"Module"`.
    // Vybe's `Display` impl on `ObjectKind::ModuleNamespace` renders
    // the value as `"[object Module]"`, which is what string coercion
    // produces.
    let out = run_js_with_imports(r#"
import * as cli from "wasi:cli";
console.log(String(cli));
"#);
    assert_eq!(out, vec!["[object Module]"]);
}

// ── Phase 6: Adapter modules (node:*) re-export from Synthetic ─────

#[test]
fn node_os_platform_returns_node_faithful_string() {
    // `node:os` is now a real host module (not a JS adapter that
    // re-exported `wasi:cli` shims). It returns Node's canonical
    // platform names — "darwin"/"linux"/"win32" — not the older
    // Vybe-shim names ("macos"/"linux"/"windows"). See
    // `crates/vybe_host/src/node/os.rs`.
    let out = run_js_with_imports(r#"
import { platform } from "node:os";
let p = platform();
console.log(p === "darwin" || p === "linux" || p === "win32");
"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn node_crypto_adapter_reexports_sha256() {
    // `node:crypto` adapter re-exports `sha256` from `vybe:crypto`.
    let out = run_js_with_imports(r#"
import { sha256 } from "node:crypto";
console.log(sha256("hello").length);
"#);
    // SHA-256 hex digest is 64 chars.
    assert_eq!(out, vec!["64"]);
}

// ── Phase 8: Link-time import validation ────────────────────────────

#[test]
fn validator_accepts_resolvable_imports() {
    // `wasi:cli/log` resolves through the registered Synthetic module.
    // Validator should return an empty unresolved list.
    use vybe_bytecode::VM;
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    vybex::adapters::register_all(&mut vm).expect("adapters");

    let module = vybe_compiler::languages::js::parse(r#"
import { log } from "wasi:cli";
log("hi");
"#).expect("parse");
    let profile = vybe_compiler::profile::parse_profile(vybe_compiler::languages::js::profile_source())
        .expect("profile");
    let module_exports = vybe_compiler::bundle::flatten_module_exports(&vm.modules);
    let result = vybe_compiler::compiler::Compiler::with_profile(profile)
        .with_module_exports(module_exports)
        .compile_with_imports(&module).expect("compile");
    let unresolved = vybe_compiler::bundle::validate_imports_against_modules(
        &result.chunks,
        &vm.modules,
    );
    assert!(
        unresolved.is_empty(),
        "expected all imports resolvable, got: {:?}",
        unresolved,
    );
}

#[test]
fn validator_flags_unknown_export() {
    // Importing a name that doesn't exist in `wasi:cli` should surface
    // at link time, not at VM setup.
    use vybe_bytecode::VM;
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    vybex::adapters::register_all(&mut vm).expect("adapters");

    let module = vybe_compiler::languages::js::parse(r#"
import { definitelyNotAThing } from "wasi:cli";
definitelyNotAThing();
"#).expect("parse");
    let profile = vybe_compiler::profile::parse_profile(vybe_compiler::languages::js::profile_source())
        .expect("profile");
    let module_exports = vybe_compiler::bundle::flatten_module_exports(&vm.modules);
    let result = vybe_compiler::compiler::Compiler::with_profile(profile)
        .with_module_exports(module_exports)
        .compile_with_imports(&module).expect("compile");
    let unresolved = vybe_compiler::bundle::validate_imports_against_modules(
        &result.chunks,
        &vm.modules,
    );
    assert!(
        unresolved.iter().any(|u| u.contains("definitelyNotAThing")),
        "expected unresolved list to include the missing export, got: {:?}",
        unresolved,
    );
}
