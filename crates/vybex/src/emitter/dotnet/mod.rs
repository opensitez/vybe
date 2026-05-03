//! Shared .NET BCL frontend for all .NET-shaped compilers (VB, C#, F#, …).
//!
//! The .NET Base Class Library exposes the same namespace hierarchy regardless
//! of language: `System.Threading.Thread.Sleep`, `System.Diagnostics.Stopwatch`,
//! etc. This module is a single source of truth so every .NET compiler resolves
//! these identically.
//!
//! ## Module structure
//!
//! - **`resolver`** — the dotted-name resolution algorithm. Owns
//!   `DottedResolution`, `ResolutionContext`, `resolve_dotted_name`,
//!   `resolve_interface_call`, and the private import-prefix matching helpers.
//!
//! - **`imports`** — the implicit import list. `default_interface_imports()`
//!   returns the namespaces every .NET compiler auto-recognises (`System`,
//!   `System.Threading`, `System.Windows.Forms`, …). Language compilers
//!   `.extend()` it with their own additions.
//!
//! - **`namespaces`** — namespace-root recognition. `is_namespace_root` /
//!   `namespace_roots` answer "is `Math` a variable or the start of a
//!   namespace chain?".
//!
//! - **`host_map`** — `.NET → Vybe host` translation tables.
//!   `namespace_to_host_module` maps `system.console` → `wasi:cli`, and
//!   `map_host_func` maps `(wasi:cli, writeline)` → `log`.
//!
//! - **`types`** — type-related lookups & predicates: `known_types()`
//!   constructor table, `is_noop_method`, `is_known_constant`, and the
//!   PascalCase name-shape helpers.
//!
//! - **`core`** — shared `.NET` library metadata that multiple framework
//!   adapters can reuse.
//!
//! - **`winforms`** — the current GUI/framework adapter, including the
//!   wrapper-class hierarchy that used to live directly under `dotnet/classes`.
//!
//! ## Future extensions
//!
//! When adding a real `Form` base class (and `Control` / `Button` / `TextBox`
//! as a class hierarchy that user code can `Inherits` from), add a new
//! `forms.rs` (and friends) sibling to these files. They use
//! `compiler_common::gui` helpers under the hood and get registered in the
//! type registry at host startup.
//!
//! Future framework frontends (MAUI, Flutter, Tkinter) follow the same
//! pattern: a sibling top-level module to `dotnet/`, structured the same way,
//! all delegating to `compiler_common::gui` for the canonical GUI emit.

pub mod resolver;
pub mod imports;
pub mod namespaces;
pub mod host_map;
pub mod types;
pub mod class_exports;
pub mod component_classes;
mod descriptor;
pub mod core;
pub mod winforms;

// ─── Public re-exports ───────────────────────────────────────────────────────
//
// The pre-split single-file `dotnet.rs` exposed everything at
// `compiler_common::dotnet::*`. The split keeps that flat surface for callers
// (compilers, walkers, host) — they continue to write
// `common::dotnet::resolve_dotted_name`, `common::dotnet::known_types`, etc.
// without caring which submodule the item lives in.

pub use resolver::{
    DottedResolution,
    ResolutionContext,
    resolve_dotted_name,
    resolve_interface_call,
};

pub use core::dotnet_core_component_descriptor;
pub use winforms::dotnet_winforms_component_descriptor;
pub use winforms::classes;

use std::sync::LazyLock;
use std::collections::HashSet;
use vybe_bytecode::component_model::{ComponentDescriptor, ComponentItemKind, ConstructorTarget, MethodBody};

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StaticMethodTarget {
    Host { module: String, func: String },
    Common { emit: String },
}

/// What an instance-method call resolves to in the Component Model
/// dispatch path. Mirrors `StaticMethodTarget` plus the `arity` so the
/// caller can validate `args.len()` against the declared signature.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum InstanceMethodTarget {
    /// Method body is a host import (typically an `ecma:*` primitive
    /// the .NET adapter delegates to).
    Host { module: String, func: String, arity: u8 },
    /// Method body is a `compiler_common` emit name routed through
    /// `dispatch::emit_common`.
    Common { emit: String, arity: u8 },
}

pub struct DotnetSurface {
    default_imports: Vec<String>,
    namespace_roots: HashSet<String>,
    noop_methods: HashSet<String>,
    known_constants: HashSet<String>,
    runtime_collection_methods: HashSet<String>,
    component_descriptor: ComponentDescriptor,
}

static DOTNET_SURFACE_CACHE: LazyLock<DotnetSurface> = LazyLock::new(build_dotnet_surface);

fn build_dotnet_surface() -> DotnetSurface {
    let component_descriptor = dotnet_component_descriptor();
    DotnetSurface {
        default_imports: imports::default_interface_imports(),
        namespace_roots: namespaces::namespace_roots(),
        noop_methods: winforms::noop_methods()
            .iter()
            .map(|name| (*name).to_string())
            .collect(),
        known_constants: core::known_constants()
            .iter()
            .map(|name| (*name).to_string())
            .collect(),
        runtime_collection_methods: collection_runtime_method_names(&component_descriptor),
        component_descriptor,
    }
}

pub fn surface() -> &'static DotnetSurface {
    &DOTNET_SURFACE_CACHE
}

impl DotnetSurface {
    pub fn default_imports(&self) -> &[String] {
        &self.default_imports
    }

    pub fn namespace_roots(&self) -> &HashSet<String> {
        &self.namespace_roots
    }

    pub fn is_namespace_root(&self, name: &str) -> bool {
        self.namespace_roots.contains(&name.to_lowercase())
    }

    pub fn is_noop_method(&self, name: &str) -> bool {
        self.noop_methods.contains(&name.to_lowercase())
    }

    pub fn is_known_constant(&self, name: &str) -> bool {
        self.known_constants.contains(&name.to_lowercase())
    }

    pub fn uses_runtime_collection_dispatch(&self, name: &str) -> bool {
        self.runtime_collection_methods.contains(&name.to_lowercase())
    }

    pub fn component_descriptor(&self) -> &ComponentDescriptor {
        &self.component_descriptor
    }

    pub fn lookup_constructor(&self, name: &str) -> Option<ConstructorTarget> {
        self.component_descriptor
            .classes
            .iter()
            .find(|class| class.name.eq_ignore_ascii_case(name))
            .and_then(|class| class.constructor.as_ref())
            .and_then(|ctor| ctor.backing.clone())
    }

    /// Component Model instance-method dispatch.
    ///
    /// Given a receiver class name (e.g. `"Dictionary"`) and a method
    /// name (e.g. `"Add"`), look up the class in the component descriptor
    /// and return its `MethodBody`. The .NET-name → ECMA-name translation
    /// happens at the descriptor level (e.g.
    /// `Dictionary.Add → MethodBody::HostCall(ecma:map.set)`), so the
    /// resulting `InstanceMethodTarget` is already in spec terms.
    ///
    /// The compiler calls this when a method-call site has a known
    /// receiver type (typically from `Dim x As New Y(...)` or
    /// `var x : Y` annotations propagated through the local table).
    /// Returns `None` for unknown classes or non-instance methods —
    /// the caller falls through to runtime dispatch (TypeRegistry hint
    /// + `__type` fallback per the compilation-hints proposal).
    pub fn lookup_instance_method(&self, class_name: &str, method_name: &str) -> Option<InstanceMethodTarget> {
        self.component_descriptor
            .classes
            .iter()
            .find(|class| class.name.eq_ignore_ascii_case(class_name))
            .and_then(|class| {
                class.methods.iter().find_map(|method| {
                    if method.is_static || !method.name.eq_ignore_ascii_case(method_name) {
                        return None;
                    }
                    match &method.body {
                        MethodBody::HostCall(target) => Some(InstanceMethodTarget::Host {
                            module: target.module.clone(),
                            func: target.name.clone(),
                            arity: method.arity,
                        }),
                        MethodBody::Common(name) => Some(InstanceMethodTarget::Common {
                            emit: name.clone(),
                            arity: method.arity,
                        }),
                        // UserChunk paths are compiled by the wrapper builder
                        // (DotnetClass) — not driven through this lookup.
                        _ => None,
                    }
                })
            })
    }

    pub fn lookup_static_method(&self, prefix: &str, method_parts: &[&str]) -> Option<StaticMethodTarget> {
        let (interface_name, type_name, method_name) = match method_parts {
            [method_name] if prefix.eq_ignore_ascii_case("application") => {
                ("system.windows.forms".to_string(), "application".to_string(), *method_name)
            }
            [method_name] => {
                let mut collected: Vec<&str> = prefix.split('.').collect();
                let type_name = collected.pop()?;
                (collected.join("."), type_name.to_string(), *method_name)
            }
            [type_name, method_name] => (prefix.to_string(), (*type_name).to_string(), *method_name),
            _ => return None,
        };

        self.component_descriptor
            .exports
            .iter()
            .find_map(|export| {
                let ComponentItemKind::Class(class) = &export.kind else {
                    return None;
                };
                let export_interface = export
                    .interface
                    .strip_prefix("dotnet.")
                    .unwrap_or(export.interface.as_str())
                    .to_lowercase();
                if export_interface != interface_name || !class.name.eq_ignore_ascii_case(&type_name) {
                    return None;
                }
                class.methods.iter().find_map(|method| {
                    if !method.is_static || !method.name.eq_ignore_ascii_case(method_name) {
                        return None;
                    }
                    match &method.body {
                        MethodBody::HostCall(target) => Some(StaticMethodTarget::Host {
                            module: target.module.clone(),
                            func: target.name.clone(),
                        }),
                        MethodBody::Common(name) => Some(StaticMethodTarget::Common { emit: name.clone() }),
                        _ => None,
                    }
                })
            })
    }
}

pub fn default_interface_imports() -> Vec<String> {
    surface().default_imports().to_vec()
}

pub fn namespace_roots() -> HashSet<String> {
    surface().namespace_roots().clone()
}

pub fn is_namespace_root(name: &str) -> bool {
    surface().is_namespace_root(name)
}

pub fn is_noop_method(name: &str) -> bool {
    surface().is_noop_method(name)
}

pub fn is_known_constant(name: &str) -> bool {
    surface().is_known_constant(name)
}

pub fn uses_runtime_collection_dispatch(name: &str) -> bool {
    surface().uses_runtime_collection_dispatch(name)
}

pub fn namespace_to_host_module(prefix: &str) -> &str {
    host_map::namespace_to_host_module(prefix)
}

pub fn map_host_func(module: &str, func: &str) -> String {
    host_map::map_host_func(module, func)
}

pub fn static_method_mappings() -> &'static [host_map::DotnetStaticMethodMapping] {
    host_map::static_method_mappings()
}

pub fn known_type_mappings() -> &'static [types::KnownTypeMapping] {
    types::known_type_mappings()
}

pub fn lookup_known_type(name: &str) -> Option<&'static types::KnownTypeMapping> {
    types::lookup_known_type(name)
}

pub fn known_types() -> std::collections::HashMap<String, types::KnownTypeTarget> {
    types::known_types()
}

pub fn capitalize_control_name(name: &str) -> String {
    types::capitalize_control_name(name)
}

pub fn capitalize_data_type(name: &str) -> String {
    types::capitalize_data_type(name)
}

/// Build the typed `.NET` core library descriptor.
///
/// This covers the shared non-WinForms surface that can be reused by multiple
/// .NET-shaped framework adapters.
/// Build the typed `.NET` framework descriptor as a compatibility merge.
///
/// Existing callers still see a single actor, but the metadata is now
/// partitioned so future framework adapters like MAUI can load separately.
pub fn dotnet_component_descriptor() -> ComponentDescriptor {
    let mut merged = dotnet_core_component_descriptor();
    descriptor::merge_component_descriptor(&mut merged, dotnet_winforms_component_descriptor());
    merged.name = "dotnet".to_string();
    merged
}

pub fn lookup_component_constructor(name: &str) -> Option<ConstructorTarget> {
    surface().lookup_constructor(name)
}

pub fn lookup_component_static_method(prefix: &str, method_parts: &[&str]) -> Option<StaticMethodTarget> {
    surface().lookup_static_method(prefix, method_parts)
}

fn collection_runtime_method_names(descriptor: &ComponentDescriptor) -> HashSet<String> {
    descriptor
        .exports
        .iter()
        .filter_map(|export| {
            if !export.interface.starts_with("dotnet.System.Collections") {
                return None;
            }
            let ComponentItemKind::Class(class) = &export.kind else {
                return None;
            };
            Some(class)
        })
        .flat_map(|class| class.methods.iter())
        .filter(|method| !method.is_static)
        .map(|method| method.name.to_lowercase())
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashSet;

    #[test]
    fn test_dotnet_component_descriptor_exports_wrapped_classes() {
        let descriptor = dotnet_component_descriptor();
        let mut expected_exports = HashSet::new();
        for export in class_exports::dotnet_class_exports() {
            expected_exports.insert(descriptor::class_export_key(export.interface, &export.class.name));
        }

        assert_eq!(descriptor.classes.len(), expected_exports.len());
        assert_eq!(descriptor.exports.len(), expected_exports.len());
        assert!(descriptor
            .exports
            .iter()
            .any(|exp| exp.interface == "dotnet.System.Windows.Forms" && exp.name == "Form"));
        assert!(descriptor
            .exports
            .iter()
            .any(|exp| exp.interface == "dotnet.System.Drawing" && exp.name == "Graphics"));
        assert!(descriptor
            .exports
            .iter()
            .any(|exp| exp.interface == "dotnet.System.Text" && exp.name == "StringBuilder"));
        assert!(descriptor
            .exports
            .iter()
            .any(|exp| exp.interface == "dotnet.System" && exp.name == "Console"));
        assert!(descriptor
            .imports
            .iter()
            .any(|imp| imp.interface == "vybe:gui" && imp.name == crate::emitter::gui::HOST_FN_SET_PROPERTY));
        assert!(descriptor
            .imports
            .iter()
            .any(|imp| imp.interface == "vybe:gui" && imp.name == "new_Form"));
        // StringBuilder no longer imports `vybe:types/stringBuilderNew`;
        // the constructor is a Common emit (`dotnet.string_builder_new`)
        // composing existing primitives. Verify the descriptor lists the
        // class export instead.
        assert!(descriptor
            .classes
            .iter()
            .any(|class| class.name == "StringBuilder"));
        let console = descriptor
            .classes
            .iter()
            .find(|class| class.name == "Console")
            .expect("Console class export");
        assert!(console.methods.iter().any(|method| method.is_static && method.name == "WriteLine"));
    }

    #[test]
    fn test_dotnet_core_component_descriptor_excludes_winforms_surface() {
        let descriptor = dotnet_core_component_descriptor();

        assert!(descriptor
            .exports
            .iter()
            .any(|exp| exp.interface == "dotnet.System" && exp.name == "Console"));
        assert!(descriptor
            .exports
            .iter()
            .any(|exp| exp.interface == "dotnet.System.Text" && exp.name == "StringBuilder"));
        assert!(!descriptor
            .exports
            .iter()
            .any(|exp| exp.interface == "dotnet.System.Windows.Forms" && exp.name == "Form"));
        assert!(!descriptor
            .imports
            .iter()
            .any(|imp| imp.interface == "vybe:gui" && imp.name == crate::emitter::gui::HOST_FN_RUN_APPLICATION));
    }

    #[test]
    fn test_dotnet_winforms_component_descriptor_contains_framework_surface() {
        let descriptor = dotnet_winforms_component_descriptor();

        assert!(descriptor
            .exports
            .iter()
            .any(|exp| exp.interface == "dotnet.System.Windows.Forms" && exp.name == "Form"));
        assert!(descriptor
            .exports
            .iter()
            .any(|exp| exp.interface == "dotnet.System.Windows.Forms" && exp.name == "Application"));
        assert!(descriptor
            .imports
            .iter()
            .any(|imp| imp.interface == "vybe:gui" && imp.name == crate::emitter::gui::HOST_FN_RUN_APPLICATION));
        assert!(!descriptor
            .exports
            .iter()
            .any(|exp| exp.interface == "dotnet.System" && exp.name == "Console"));
    }

    #[test]
    fn test_lookup_component_constructor_uses_descriptor_surface() {
        // `StringBuilder` materializes via the shared `dotnet.string_builder_new`
        // adapter (plain Object + `__buffer` string), not a `vybe:types`
        // host fn. The Common-emit path keeps the construction logic in
        // one Rust file (`emitter/dotnet/core/stringbuilder_adapter.rs`).
        let binding = lookup_component_constructor("StringBuilder").expect("StringBuilder constructor");
        assert_eq!(binding, ConstructorTarget::Common("dotnet.string_builder_new".to_string()));
    }

    #[test]
    fn test_lookup_component_constructor_supports_common_emit() {
        let binding = lookup_component_constructor("List").expect("List constructor");
        assert_eq!(binding, ConstructorTarget::Common("collections.new".to_string()));
    }

    #[test]
    fn test_lookup_component_static_method_uses_descriptor_surface() {
        // `Console.WriteLine` routes through the shared `dotnet.console_writeline`
        // adapter (capitalises bools, maps null→"") rather than calling
        // `wasi:cli.log` directly. C# / VB / any future .NET-shape language
        // gets the same behaviour by listing the same `Common` emit target.
        let binding = lookup_component_static_method("system.console", &["writeline"])
            .expect("Console.WriteLine static method");
        assert_eq!(binding, StaticMethodTarget::Common {
            emit: "dotnet.console_writeline".to_string(),
        });
    }

    #[test]
    fn test_runtime_collection_dispatch_uses_descriptor_surface() {
        assert!(uses_runtime_collection_dispatch("Add"));
        assert!(uses_runtime_collection_dispatch("ContainsKey"));
        assert!(uses_runtime_collection_dispatch("Clear"));
        assert!(!uses_runtime_collection_dispatch("WriteLine"));
    }

    #[test]
    fn test_dotnet_surface_cache_merges_metadata_predicates() {
        assert!(default_interface_imports().contains(&"system.windows.forms".to_string()));
        assert!(is_namespace_root("application"));
        assert!(is_noop_method("SuspendLayout"));
        assert!(is_known_constant("PI"));
    }
}
