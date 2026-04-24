//! ECMA-262 §16.2.1 Abstract Module Records — the canonical registry for
//! host modules (`wasi:*`, `wasm:js-*`, `vybe:*`), `.wasm` source modules,
//! and in-language adapter modules (`node:http` authored as a `.js` file
//! that re-exports from `wasi:http`).
//!
//! Phase 1 of the ESM host-access migration plan: this module introduces
//! the data shape. The VM populates `vm.modules` alongside the existing
//! flat `host_registry` on every `register_host_fn` call — no behavior
//! change. Later phases build the Linker on top of this registry.
//!
//! See `esmhostplan.md` at the project root for the full migration plan.

use std::collections::{BTreeMap, HashMap};
use crate::value::Value;

/// A registered module — Synthetic (Rust-backed), Wasm (loaded `.wasm`),
/// or Adapter (source-language file re-exporting from Synthetic modules).
#[derive(Debug)]
pub struct ModuleRecord {
    /// Canonical specifier: `"wasi:cli/environment"`, `"vybe:js-math"`,
    /// `"./helpers.wasm"`, etc. Package ≠ interface — this is the
    /// *interface* path, which is what source code imports from.
    pub specifier: String,
    pub kind: ModuleKind,
    pub status: ModuleStatus,
    /// Name → binding. For Synthetic modules each host fn registration
    /// inserts a `Function` entry. Adapter modules carry `Indirect`
    /// entries pointing at their source.
    pub exports: HashMap<String, ExportEntry>,
    /// §16.2.1.3 — direct dependencies with their import attributes.
    /// Empty for Synthetic leaves; populated for Wasm / Adapter once
    /// their module graph is walked at link time (later phases).
    pub requested_modules: Vec<ModuleRequest>,
    /// Capability required to link this module. `None` = unrestricted
    /// (e.g. `wasi:cli/log`). `Some("filesystem")` means the Linker
    /// must check the active `Capabilities` set and fail the link if
    /// denied. Populated later; Phase 1 leaves it `None` for every
    /// auto-upgraded registration.
    pub capability: Option<String>,
}

impl ModuleRecord {
    pub fn new_synthetic(specifier: impl Into<String>) -> Self {
        Self {
            specifier: specifier.into(),
            kind: ModuleKind::Synthetic,
            // Synthetic modules have no source code to evaluate — they
            // are linked immediately on registration.
            status: ModuleStatus::Linked,
            exports: HashMap::new(),
            requested_modules: Vec::new(),
            capability: None,
        }
    }
}

/// What kind of module this is. Governs how the Linker and Evaluate
/// steps behave.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ModuleKind {
    /// `wasi:*`, `wasm:js-*`, `vybe:*` — backed by Rust closures in
    /// `vybe_host`. No source code, no evaluation. Linked on
    /// registration.
    Synthetic,
    /// `.wasm` loaded via the ESM Integration proposal.
    Wasm,
    /// User-authored source-language module (e.g. `node:http.js`)
    /// whose only job is to re-export from Synthetic modules.
    Adapter,
}

/// §16.2.1.4 — the module lifecycle state machine. Synthetic modules
/// jump straight to `Linked`; Wasm / Adapter walk the full path.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ModuleStatus {
    New,
    Linking,
    Linked,
    Evaluating,
    Evaluated,
    Errored,
}

/// A single named export from a module.
#[derive(Debug, Clone)]
pub enum ExportEntry {
    /// Callable — `host_fns[idx]` for Synthetic, or a function ref
    /// index for Wasm. Compiled as `CALL_IMPORT` when the import is
    /// called, or as a `Value::Object(HostFunction(idx))` when read
    /// as a value.
    Function { idx: usize },
    /// Immutable value export — `vybe:js-math.PI`, etc. Compiled as
    /// `emit_const`.
    Value(Value),
    /// Component Model resource type (WIT `resource`). Registered in
    /// the `TypeRegistry`; `new SomeResource(...)` dispatches via
    /// the type id.
    ResourceType { type_id: usize },
    /// Re-export: `export { X } from "other"`. The Linker resolves
    /// transitively so the importer binds to the final target
    /// directly — no runtime chase.
    Indirect { from: String, name: String },
}

/// §16.2.1.3 — `(specifier, attributes)` identifies the requested
/// module. Import attributes (`with { type: "json" }`) are reserved
/// in the shape; today only `{}` is observable.
#[derive(Debug, Clone, Default)]
pub struct ModuleRequest {
    pub specifier: String,
    pub attributes: BTreeMap<String, String>,
}
