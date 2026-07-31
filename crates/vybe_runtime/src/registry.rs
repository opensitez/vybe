//! Language plugin registry — the seam that lets `vybe_compiler` reach a
//! language without naming its module. Each language registers a full
//! `LanguageDef` (function pointers for parse / profile / emit-dispatch /
//! class-normalisation / namespace-tree); the compiler looks them up by name.
//!
//! This is what severs `compiler → languages::<lang>`, and — when a language
//! becomes a loadable dylib — it is exactly the entry point the host calls
//! after `dlopen` to hand back the plugin.

use std::sync::{Mutex, OnceLock};

use vybe_ast::{ClassMember, ClassModifiers, Module, Span};
use vybe_ast::class_normalize::NormalClass;
use crate::Chunk;

/// Parse source → common AST.
pub type ParseFn = fn(&str) -> Result<Module, String>;
/// Emit a `common:<lang>.*` op inline. Returns `false` if unhandled.
pub type EmitDispatchFn = fn(&str, &mut Vec<Chunk>, usize, u8, u32) -> bool;
/// Per-language class normalisation (walker → language-agnostic `NormalClass`).
pub type NormalizeClassFn = fn(
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> NormalClass;

/// A registered language's full compiler-facing surface. All fields are plain
/// function pointers (+ a `&'static str`), so the struct is `Copy` and crosses
/// the dylib boundary cleanly. `Option` fields are capabilities a language may
/// lack (e.g. `c` has a namespace tree but no class normalisation; `csharp`
/// has no own emit-dispatch — it rides the dotnet platform).
#[derive(Clone, Copy)]
pub struct LanguageDef {
    pub name: &'static str,
    pub parse: ParseFn,
    pub profile_source: fn() -> &'static str,
    pub emit_dispatch: Option<EmitDispatchFn>,
    pub normalize_class: Option<NormalizeClassFn>,
    pub register_tree: Option<fn()>,
}

fn registry() -> &'static Mutex<Vec<LanguageDef>> {
    static REGISTRY: OnceLock<Mutex<Vec<LanguageDef>>> = OnceLock::new();
    REGISTRY.get_or_init(|| Mutex::new(Vec::new()))
}

/// Register a language. Idempotent by `name`, so a language crate can safely
/// call this from its own initialiser (or its dylib entry point).
pub fn register_language(def: LanguageDef) {
    let mut r = registry().lock().unwrap();
    if !r.iter().any(|p| p.name == def.name) {
        r.push(def);
    }
}

/// The plugin registered under `name`, if any.
pub fn find(name: &str) -> Option<LanguageDef> {
    registry()
        .lock()
        .unwrap()
        .iter()
        .find(|p| p.name == name)
        .copied()
}

/// All registered plugins (a snapshot).
pub fn all() -> Vec<LanguageDef> {
    registry().lock().unwrap().clone()
}

/// The emit-dispatcher that owns the `common:<prefix>.*` namespace, if a
/// language registered one.
pub fn emit_dispatch_for(prefix: &str) -> Option<EmitDispatchFn> {
    find(prefix).and_then(|p| p.emit_dispatch)
}

/// Dispatch class normalisation to the registered language, or `None` if the
/// language has no normaliser.
pub fn normalize_class(
    lang: &str,
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> Option<NormalClass> {
    let f = find(lang).and_then(|p| p.normalize_class)?;
    Some(f(span, name, parents, interfaces, members, modifiers))
}

// ── Platform plugins ────────────────────────────────────────────────────────
//
// Platforms (dotnet, flutter, libc, plib, wasm) are plugins in exactly the same
// sense as languages: they contribute a namespace tree and an emit dispatcher,
// and the compiler must not name them. They were compile-time dependencies of
// `vybe_compiler`, which is legacy — the whole point of this seam is that a
// plugin can become a dylib loaded at run time, and a `Cargo.toml` edge makes
// that impossible.
//
// Same shape as `LanguageDef`: plain function pointers, `Copy`, crosses a dylib
// boundary cleanly.

/// A registered platform's compiler-facing surface.
#[derive(Clone, Copy)]
pub struct PlatformDef {
    /// The `common:<name>.*` prefix this platform owns (`"dotnet"`, `"libc"`).
    pub name: &'static str,
    /// Emit a `common:<name>.*` op inline. `None` for a platform that only
    /// contributes a namespace tree.
    pub emit_dispatch: Option<EmitDispatchFn>,
    /// Mount this platform's namespace tree.
    pub register_tree: Option<fn()>,

    // ── Data the compiler used to reach by naming the crate ──────────────
    //
    // These were direct `crate::platforms::dotnet::…` calls, which is why
    // `vybe_compiler` had a Cargo dependency on the platform crates at all —
    // registration was never the reason. Every signature already uses shared
    // types (`bool`, `Chunk`, `ComponentDescriptor` from `component_model`), so
    // they are plain function pointers and cross a dylib boundary cleanly.
    /// Namespace constants this platform contributes (`Math.PI` → 3.14159…).
    pub namespace_constants: Option<fn() -> &'static [(&'static str, f64)]>,
    /// This platform's component descriptor (classes/methods it exports).
    pub component_descriptor: Option<fn() -> crate::component_model::ComponentDescriptor>,
    /// True when `name` is a class the platform's descriptor owns.
    pub is_descriptor_class: Option<fn(&str) -> bool>,
    /// Build the platform's numeric-format runtime helper chunk.
    pub numeric_format_helper: Option<fn(&mut Chunk) -> Chunk>,
    // NOTE: there are deliberately NO per-platform lookup hooks here
    // (constructor / instance-method / instance-property / known-constant /
    // …). A platform declares its classes ONCE, as `Type` nodes in the
    // namespace tree, and the compiler resolves through the tree. A second
    // function-pointer surface answering the same questions is duplication —
    // adapters over one interface, not a per-platform API.
    /// Decode a binary module this platform understands into chunks
    /// (`.wasm` → `Vec<Chunk>`). Both sides of the signature are shared types.
    pub read_binary_module: Option<fn(&[u8]) -> Result<Vec<Chunk>, String>>,
}

fn platform_registry() -> &'static Mutex<Vec<PlatformDef>> {
    static REGISTRY: OnceLock<Mutex<Vec<PlatformDef>>> = OnceLock::new();
    REGISTRY.get_or_init(|| Mutex::new(Vec::new()))
}

/// Register a platform. Idempotent by `name`, so a platform crate can call this
/// from its own initialiser or its dylib entry point.
pub fn register_platform(def: PlatformDef) {
    let mut r = platform_registry().lock().unwrap();
    if !r.iter().any(|p| p.name == def.name) {
        r.push(def);
    }
}

/// All registered platforms (a snapshot).
pub fn all_platforms() -> Vec<PlatformDef> {
    platform_registry().lock().unwrap().clone()
}

/// The emit-dispatcher owning the `common:<prefix>.*` namespace, if a platform
/// registered one. Replaces a hardcoded `match prefix { "dotnet" => …, "libc"
/// => … }` in the compiler — the same name-check antipattern as
/// `profile.name == "<lang>"`, one layer up.
pub fn platform_emit_dispatch_for(prefix: &str) -> Option<EmitDispatchFn> {
    platform_registry()
        .lock()
        .unwrap()
        .iter()
        .find(|p| p.name == prefix)
        .and_then(|p| p.emit_dispatch)
}

/// The first registered platform that answers each query. The compiler asks the
/// REGISTRY, never a named crate — that is what removes the Cargo edge and lets
/// a platform ship as a dylib.
pub fn platform_namespace_constants() -> Vec<&'static (&'static str, f64)> {
    all_platforms()
        .iter()
        .filter_map(|p| p.namespace_constants)
        .flat_map(|f| f().iter())
        .collect()
}

/// True when any registered platform's descriptor owns `name`.
pub fn platform_owns_descriptor_class(name: &str) -> bool {
    all_platforms()
        .iter()
        .filter_map(|p| p.is_descriptor_class)
        .any(|f| f(name))
}

/// Every registered platform's component descriptor.
pub fn platform_component_descriptors() -> Vec<crate::component_model::ComponentDescriptor> {
    all_platforms()
        .iter()
        .filter_map(|p| p.component_descriptor)
        .map(|f| f())
        .collect()
}

/// The numeric-format helper builder, if a platform provides one.
pub fn platform_numeric_format_helper() -> Option<fn(&mut Chunk) -> Chunk> {
    all_platforms()
        .iter()
        .find_map(|p| p.numeric_format_helper)
}




/// Decode a binary module through whichever platform can read it.
pub fn platform_read_binary_module(data: &[u8]) -> Option<Result<Vec<Chunk>, String>> {
    all_platforms()
        .iter()
        .find_map(|p| p.read_binary_module)
        .map(|f| f(data))
}

/// Mount every registered platform's namespace tree.
pub fn register_all_platform_trees() {
    let fns: Vec<fn()> = all_platforms()
        .iter()
        .filter_map(|p| p.register_tree)
        .collect();
    for f in fns {
        f();
    }
}

/// Mount every registered language's namespace tree.
pub fn register_all_trees() {
    let fns: Vec<fn()> = all().iter().filter_map(|p| p.register_tree).collect();
    for f in fns {
        f();
    }
}

// ── Optional language hooks ─────────────────────────────────────────────────
//
// Language-specific compiler behaviours the core calls *if present*. Each is
// `Option` — a language registers only what it implements (relational compare,
// JS proxy dispatch, PHP autoload/source-normalise, Python value-eq, …).

/// How a language spells names that live in a VARIABLE namespace kept separate
/// from its function/class namespace — in PHP, `$x` and `x()` are unrelated
/// bindings that can coexist. Most languages have ONE namespace, so most leave
/// [`LanguageHooks::variable_namespace`] `None` and every name is just a name.
///
/// The marker carrying that distinction (`$`, `@`, a sigil, a prefix) is the
/// language's own spelling, so the language owns all of it. These three
/// operations are ONE mechanism and must agree with each other: `global_key`
/// consumes the canonicalized output of `body`, and `spell` is `body`'s
/// inverse. Changing one alone breaks the set.
pub struct VariableNamespace {
    /// `"$x"` → `"x"`: the marker removed, so canonicalization folds the bare
    /// name. MUST return its input unchanged for a name outside the namespace —
    /// the core uses "did this change the name" as the is-a-variable test.
    pub body: fn(&str) -> &str,
    /// `"x"` → `"$x"`: the inverse of [`Self::body`], for building a reference
    /// to a variable the compiler knows only by its bare name (PHP `compact`,
    /// `extract`).
    pub spell: fn(&str) -> String,
    /// The key this variable takes at GLOBAL scope, given the source name and
    /// its canonicalized body. This is what keeps a global `$foo` from
    /// colliding with a function `foo`. `None` = use the canonical body
    /// unchanged, which is how names the host registers under their literal
    /// source spelling (PHP superglobals) stay reachable.
    pub global_key: fn(&str, &str) -> Option<String>,
}

#[derive(Clone, Copy, Default)]
pub struct LanguageHooks {
    pub value_eq: Option<fn(&mut Chunk, u32)>,
    /// See [`VariableNamespace`]. `None` (the default) = one namespace.
    pub variable_namespace: Option<&'static VariableNamespace>,
    pub relational_compare: Option<fn(&mut Chunk, fn(&mut Chunk, u32), u32)>,
    /// `+` where the language overloads it on a COLLECTION as well as on
    /// numbers — PHP's array union. Same shape as `relational_compare`: the
    /// core cannot decide this from the operator alone, and the decision is a
    /// property of the language's `+`, not of any operand the compiler can see.
    /// Stack: `[l, r] → [result]`.
    pub arith_add: Option<fn(&mut Vec<Chunk>, usize, u32)>,
    /// How ONE operand of a string-concatenation operator becomes a string,
    /// for the languages whose concat coerces both sides up front
    /// (`profile.concat_stringifies_operands`). Only needed where that
    /// spelling differs from the shared `to_string` — PHP renders `true` as
    /// `"1"`, `null` as `""`, an array as `"Array"`. Leave it `None` to get
    /// the shared coercion. Stack: `[value] → [string]`.
    pub concat_stringify: Option<fn(&mut Vec<Chunk>, usize, u32)>,
    pub constructor_ref_autoload: Option<fn(&mut Chunk, &str, &str, u32)>,
    pub dynamic_constructor_ref_autoload: Option<fn(&mut Chunk, &str, Option<&str>, &str, u32)>,
    pub proxy_get: Option<fn(&mut [Chunk], usize, u32)>,
    pub proxy_set: Option<fn(&mut [Chunk], usize, u32)>,
    pub proxy_set_bool: Option<fn(&mut [Chunk], usize, u32)>,
    pub proxy_has: Option<fn(&mut [Chunk], usize, u32)>,
    pub proxy_create: Option<fn(&mut [Chunk], usize, u32)>,
    pub normalize_source: Option<fn(&str) -> String>,
    pub str_getcsv: Option<fn(&mut [Chunk], usize, u8, u32)>,
    /// Parse an eval/`Function`-body string in *its own* source coordinates
    /// (no language prelude prepended), so completion-value and span logic in
    /// the dynamic runtime stays correct. JS registers this; the runtime looks
    /// it up by name instead of naming the language crate directly.
    pub parse_eval: Option<ParseFn>,
}

fn hooks_registry() -> &'static Mutex<Vec<(&'static str, LanguageHooks)>> {
    static H: OnceLock<Mutex<Vec<(&'static str, LanguageHooks)>>> = OnceLock::new();
    H.get_or_init(|| Mutex::new(Vec::new()))
}

/// Register a language's optional hooks (idempotent by name).
pub fn register_hooks(name: &'static str, hooks: LanguageHooks) {
    let mut r = hooks_registry().lock().unwrap();
    if !r.iter().any(|(n, _)| *n == name) {
        r.push((name, hooks));
    }
}

/// The hooks registered for `name` (all-`None` default if none registered).
pub fn hooks(name: &str) -> LanguageHooks {
    hooks_registry()
        .lock()
        .unwrap()
        .iter()
        .find(|(n, _)| *n == name)
        .map(|(_, h)| *h)
        .unwrap_or_default()
}
