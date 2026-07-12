//! Language plugin registry — the seam that lets `vybe_compiler` reach a
//! language without naming its module. Each language registers a full
//! `LanguagePlugin` (function pointers for parse / profile / emit-dispatch /
//! class-normalisation / namespace-tree); the compiler looks them up by name.
//!
//! This is what severs `compiler → languages::<lang>`, and — when a language
//! becomes a loadable dylib — it is exactly the entry point the host calls
//! after `dlopen` to hand back the plugin.

use std::sync::{Mutex, OnceLock};

use vybe_ast::{ClassMember, ClassModifiers, Module, Span};
use vybe_bytecode::Chunk;

use crate::class_normalize::NormalClass;

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
pub struct LanguagePlugin {
    pub name: &'static str,
    pub parse: ParseFn,
    pub profile_source: fn() -> &'static str,
    pub emit_dispatch: Option<EmitDispatchFn>,
    pub normalize_class: Option<NormalizeClassFn>,
    pub register_tree: Option<fn()>,
}

fn registry() -> &'static Mutex<Vec<LanguagePlugin>> {
    static REGISTRY: OnceLock<Mutex<Vec<LanguagePlugin>>> = OnceLock::new();
    REGISTRY.get_or_init(|| Mutex::new(Vec::new()))
}

/// Register a language. Idempotent by `name`, so a language crate can safely
/// call this from its own initialiser (or its dylib entry point).
pub fn register_language(plugin: LanguagePlugin) {
    let mut r = registry().lock().unwrap();
    if !r.iter().any(|p| p.name == plugin.name) {
        r.push(plugin);
    }
}

/// The plugin registered under `name`, if any.
pub fn find(name: &str) -> Option<LanguagePlugin> {
    registry().lock().unwrap().iter().find(|p| p.name == name).copied()
}

/// All registered plugins (a snapshot).
pub fn all() -> Vec<LanguagePlugin> {
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

#[derive(Clone, Copy, Default)]
pub struct LanguageHooks {
    pub value_eq: Option<fn(&mut Chunk, u32)>,
    pub relational_compare: Option<fn(&mut Chunk, fn(&mut Chunk, u32), u32)>,
    pub constructor_ref_autoload: Option<fn(&mut Chunk, &str, &str, u32)>,
    pub dynamic_constructor_ref_autoload:
        Option<fn(&mut Chunk, &str, Option<&str>, &str, u32)>,
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
