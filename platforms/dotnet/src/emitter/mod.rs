//! Shared adapter surface for all `.NET`-shaped compilers (VB, C#, F#, ...).
//!
//! Source languages can talk in `System.*`, WinForms, and other `.NET`-shaped
//! names, but this is not a `.NET` VM. This module is the compiler-side adapter
//! layer that rewrites those shapes onto the real runtime capability modules
//! available inside the wasm VM: `ecma:*`, `node:*`, `wasi:*`, `web:*`, and
//! `vybe:*`.
//!
//! In other words: VB/C# think they are targeting `.NET`; the generated wasm is
//! actually targeting JS/WASM-flavoured runtime primitives underneath.
pub mod class_exports;
pub mod dispatch;
pub mod core;
mod descriptor;
pub mod host_map;
pub mod imports;
pub mod namespaces;
pub mod resolver;
pub mod types;
pub mod winforms;
pub use core::dotnet_core_component_descriptor;
pub use resolver::{
    DottedResolution, ResolutionContext, resolve_dotted_name, resolve_interface_call,
};
use std::collections::{HashMap, HashSet};
use std::sync::LazyLock;
use vybe_bytecode::component_model::{
    ComponentDescriptor, ComponentItemKind, ConstructorTarget, MethodBody,
};
pub use winforms::classes;
pub use winforms::dotnet_winforms_component_descriptor;
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StaticMethodTarget {
    Host { module: String, func: String },
    Common { emit: String },
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StaticPropertyTarget {
    Host { module: String, func: String },
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum InstanceMethodTarget {
    Host {
        module: String,
        func: String,
        arity: u8,
    },
    Common {
        emit: String,
        arity: u8,
    },
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum InstancePropertyTarget {
    Host { module: String, func: String },
}
pub struct DotnetSurface {
    default_imports: Vec<String>,
    namespace_roots: HashSet<String>,
    noop_methods: HashSet<String>,
    known_constants: HashSet<String>,
    runtime_collection_methods: HashSet<String>,
    runtime_collection_method_arities: HashMap<String, HashSet<u8>>,
    component_descriptor: ComponentDescriptor,
}

#[cfg(test)]
fn is_real_runtime_interface(interface: &str) -> bool {
    interface.contains(':')
        && !interface.starts_with("dotnet.")
        && !interface.starts_with("system.")
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
        runtime_collection_method_arities: collection_runtime_method_arities(&component_descriptor),
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
        self.runtime_collection_methods
            .contains(&name.to_lowercase())
    }

    pub fn uses_runtime_collection_dispatch_arity(&self, name: &str, arg_count: u8) -> bool {
        self.runtime_collection_method_arities
            .get(&name.to_lowercase())
            .map(|arities| arities.contains(&arg_count))
            .unwrap_or(false)
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
    pub fn lookup_instance_method(
        &self,
        class_name: &str,
        method_name: &str,
        arg_count: u8,
    ) -> Option<InstanceMethodTarget> {
        let requested = class_name.trim();
        let requested_short = requested.rsplit('.').next().unwrap_or(requested);
        self.component_descriptor
            .classes
            .iter()
            .find(|class| {
                class.name.eq_ignore_ascii_case(requested)
                    || class.name.eq_ignore_ascii_case(requested_short)
            })
            .and_then(|class| {
                class
                    .methods
                    .iter()
                    .filter(|method| {
                        !method.is_static && method.name.eq_ignore_ascii_case(method_name)
                    })
                    .find(|method| method.arity == arg_count)
                    .or_else(|| {
                        // Backward-compatible fallback for classes that only
                        // define one method with this name.
                        class.methods.iter().find(|method| {
                            !method.is_static && method.name.eq_ignore_ascii_case(method_name)
                        })
                    })
                    .and_then(|method| {
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

    pub fn lookup_instance_method_return_type(
        &self,
        class_name: &str,
        method_name: &str,
        arg_count: u8,
    ) -> Option<String> {
        let requested = class_name.trim();
        let requested_short = requested.rsplit('.').next().unwrap_or(requested);
        self.component_descriptor
            .classes
            .iter()
            .find(|class| {
                class.name.eq_ignore_ascii_case(requested)
                    || class.name.eq_ignore_ascii_case(requested_short)
            })
            .and_then(|class| {
                class
                    .methods
                    .iter()
                    .filter(|method| {
                        !method.is_static && method.name.eq_ignore_ascii_case(method_name)
                    })
                    .find(|method| method.arity == arg_count)
                    .or_else(|| {
                        class.methods.iter().find(|method| {
                            !method.is_static && method.name.eq_ignore_ascii_case(method_name)
                        })
                    })
                    .and_then(|method| {
                        dotnet_instance_method_return_type(&class.name, &method.name)
                    })
            })
    }

    pub fn lookup_instance_property(
        &self,
        class_name: &str,
        property_name: &str,
    ) -> Option<InstancePropertyTarget> {
        let requested = class_name.trim();
        let requested_short = requested.rsplit('.').next().unwrap_or(requested);
        self.component_descriptor
            .classes
            .iter()
            .find(|class| {
                class.name.eq_ignore_ascii_case(requested)
                    || class.name.eq_ignore_ascii_case(requested_short)
            })
            .and_then(|class| {
                class
                    .properties
                    .iter()
                    .find(|property| property.name.eq_ignore_ascii_case(property_name))
            })
            .and_then(|property| {
                property
                    .getter
                    .as_ref()
                    .map(|target| InstancePropertyTarget::Host {
                        module: target.module.clone(),
                        func: target.name.clone(),
                    })
            })
    }

    pub fn lookup_static_method(
        &self,
        prefix: &str,
        method_parts: &[&str],
    ) -> Option<StaticMethodTarget> {
        let (interface_name, type_name, method_name) = match method_parts {
            [method_name] if prefix.eq_ignore_ascii_case("application") => (
                "system.windows.forms".to_string(),
                "application".to_string(),
                *method_name,
            ),
            [method_name] => {
                let mut collected: Vec<&str> = prefix.split('.').collect();
                let type_name = collected.pop()?;
                (collected.join("."), type_name.to_string(), *method_name)
            }
            [type_name, method_name] => {
                (prefix.to_string(), (*type_name).to_string(), *method_name)
            }
            _ => return None,
        };

        self.component_descriptor.exports.iter().find_map(|export| {
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
                    MethodBody::Common(name) => {
                        Some(StaticMethodTarget::Common { emit: name.clone() })
                    }
                    _ => None,
                }
            })
        })
    }

    pub fn lookup_static_property(
        &self,
        prefix: &str,
        property_name: &str,
    ) -> Option<StaticPropertyTarget> {
        let (interface_name, type_name) = if prefix.eq_ignore_ascii_case("application") {
            (
                "system.windows.forms".to_string(),
                "application".to_string(),
            )
        } else {
            let mut collected: Vec<&str> = prefix.split('.').collect();
            let type_name = collected.pop()?;
            (collected.join("."), type_name.to_string())
        };

        self.component_descriptor.exports.iter().find_map(|export| {
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
            class.properties.iter().find_map(|property| {
                if !property.name.eq_ignore_ascii_case(property_name) {
                    return None;
                }
                property
                    .getter
                    .as_ref()
                    .map(|target| StaticPropertyTarget::Host {
                        module: target.module.clone(),
                        func: target.name.clone(),
                    })
            })
        })
    }
}

fn dotnet_instance_method_return_type(class_name: &str, method_name: &str) -> Option<String> {
    let class = class_name.rsplit('.').next().unwrap_or(class_name);
    if class.eq_ignore_ascii_case("StringBuilder") {
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "append" | "appendline" | "appendformat" | "insert" | "remove" | "replace" | "clear"
        ) {
            return Some("StringBuilder".into());
        }
        if method_name.eq_ignore_ascii_case("ToString") {
            return Some("string".into());
        }
    }
    None
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

pub fn uses_runtime_collection_dispatch_arity(name: &str, arg_count: u8) -> bool {
    surface().uses_runtime_collection_dispatch_arity(name, arg_count)
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

pub fn namespace_constant_mappings() -> &'static [(&'static str, f64)] {
    winforms::types::namespace_constants()
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

pub fn lookup_component_static_method(
    prefix: &str,
    method_parts: &[&str],
) -> Option<StaticMethodTarget> {
    surface().lookup_static_method(prefix, method_parts)
}

pub fn lookup_component_static_property(
    prefix: &str,
    property_name: &str,
) -> Option<StaticPropertyTarget> {
    surface().lookup_static_property(prefix, property_name)
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

fn collection_runtime_method_arities(
    descriptor: &ComponentDescriptor,
) -> HashMap<String, HashSet<u8>> {
    let mut out: HashMap<String, HashSet<u8>> = HashMap::new();
    for export in &descriptor.exports {
        if !export.interface.starts_with("dotnet.System.Collections") {
            continue;
        }
        let ComponentItemKind::Class(class) = &export.kind else {
            continue;
        };
        for method in &class.methods {
            if method.is_static {
                continue;
            }
            out.entry(method.name.to_lowercase())
                .or_default()
                .insert(method.arity);
        }
    }
    out
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
            expected_exports.insert(descriptor::class_export_key(
                export.interface,
                &export.class.name,
            ));
        }

        assert_eq!(descriptor.classes.len(), expected_exports.len());
        assert_eq!(descriptor.exports.len(), expected_exports.len());
        assert!(
            descriptor
                .exports
                .iter()
                .any(|exp| exp.interface == "dotnet.System.Windows.Forms" && exp.name == "Form")
        );
        assert!(
            descriptor
                .exports
                .iter()
                .any(|exp| exp.interface == "dotnet.System.Drawing" && exp.name == "Graphics")
        );
        assert!(
            descriptor
                .exports
                .iter()
                .any(|exp| exp.interface == "dotnet.System.Text" && exp.name == "StringBuilder")
        );
        assert!(
            descriptor
                .exports
                .iter()
                .any(|exp| exp.interface == "dotnet.System" && exp.name == "Console")
        );
        assert!(
            descriptor
                .imports
                .iter()
                .any(|imp| imp.interface == "vybe:gui"
                    && imp.name == vybe_emitter::gui::HOST_FN_SET_PROPERTY)
        );
        assert!(
            descriptor
                .imports
                .iter()
                .any(|imp| imp.interface == "vybe:gui" && imp.name == "new_Form")
        );
        // StringBuilder no longer imports `vybe:types/stringBuilderNew`;
        // the constructor is a Common emit (`dotnet.string_builder_new`)
        // composing existing primitives. Verify the descriptor lists the
        // class export instead.
        assert!(
            descriptor
                .classes
                .iter()
                .any(|class| class.name == "StringBuilder")
        );
        let console = descriptor
            .classes
            .iter()
            .find(|class| class.name == "Console")
            .expect("Console class export");
        assert!(
            console
                .methods
                .iter()
                .any(|method| method.is_static && method.name == "WriteLine")
        );
    }

    #[test]
    fn test_dotnet_core_component_descriptor_excludes_winforms_surface() {
        let descriptor = dotnet_core_component_descriptor();

        assert!(
            descriptor
                .exports
                .iter()
                .any(|exp| exp.interface == "dotnet.System" && exp.name == "Console")
        );
        assert!(
            descriptor
                .exports
                .iter()
                .any(|exp| exp.interface == "dotnet.System.Text" && exp.name == "StringBuilder")
        );
        assert!(
            !descriptor
                .exports
                .iter()
                .any(|exp| exp.interface == "dotnet.System.Windows.Forms" && exp.name == "Form")
        );
        assert!(
            !descriptor
                .imports
                .iter()
                .any(|imp| imp.interface == "vybe:gui"
                    && imp.name == vybe_emitter::gui::HOST_FN_RUN_APPLICATION)
        );
    }

    #[test]
    fn test_dotnet_winforms_component_descriptor_contains_framework_surface() {
        let descriptor = dotnet_winforms_component_descriptor();

        assert!(
            descriptor
                .exports
                .iter()
                .any(|exp| exp.interface == "dotnet.System.Windows.Forms" && exp.name == "Form")
        );
        assert!(
            descriptor
                .exports
                .iter()
                .any(|exp| exp.interface == "dotnet.System.Windows.Forms"
                    && exp.name == "Application")
        );
        assert!(
            descriptor
                .imports
                .iter()
                .any(|imp| imp.interface == "vybe:gui"
                    && imp.name == vybe_emitter::gui::HOST_FN_RUN_APPLICATION)
        );
        assert!(
            !descriptor
                .exports
                .iter()
                .any(|exp| exp.interface == "dotnet.System" && exp.name == "Console")
        );
    }

    #[test]
    fn test_dotnet_component_descriptors_import_only_real_runtime_interfaces() {
        for descriptor in [
            dotnet_core_component_descriptor(),
            dotnet_winforms_component_descriptor(),
        ] {
            for import in &descriptor.imports {
                assert!(
                    is_real_runtime_interface(&import.interface),
                    "dotnet adapter leaked non-runtime import interface: {}::{}",
                    import.interface,
                    import.name,
                );
            }
        }
    }

    #[test]
    fn test_lookup_component_constructor_uses_descriptor_surface() {
        // `StringBuilder` materializes via the shared `dotnet.string_builder_new`
        // adapter (plain Object + `__buffer` string), not a `vybe:types`
        // host fn. The Common-emit path keeps the construction logic in
        // one Rust file (`emitter/dotnet/core/stringbuilder_adapter.rs`).
        let binding =
            lookup_component_constructor("StringBuilder").expect("StringBuilder constructor");
        assert_eq!(
            binding,
            ConstructorTarget::Common("dotnet.string_builder_new".to_string())
        );
    }

    #[test]
    fn test_lookup_component_constructor_supports_common_emit() {
        let binding = lookup_component_constructor("List").expect("List constructor");
        assert_eq!(
            binding,
            ConstructorTarget::Common("collections.new".to_string())
        );
    }

    #[test]
    fn test_lookup_component_static_method_uses_descriptor_surface() {
        // `Console.WriteLine` routes through the shared `dotnet.console_writeline`
        // adapter (capitalises bools, maps null→"") rather than calling
        // `wasi:cli.log` directly. C# / VB / any future .NET-shape language
        // gets the same behaviour by listing the same `Common` emit target.
        let binding = lookup_component_static_method("system.console", &["writeline"])
            .expect("Console.WriteLine static method");
        assert_eq!(
            binding,
            StaticMethodTarget::Common {
                emit: "dotnet.console_writeline".to_string(),
            }
        );
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
