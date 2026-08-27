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
pub mod core;
mod descriptor;
pub mod dispatch;
pub mod host_map;
pub mod imports;
pub mod tree_register;
pub mod types;
pub mod winforms;
pub use core::dotnet_core_component_descriptor;
use std::collections::{HashMap, HashSet};
use std::sync::LazyLock;
use vybe_runtime::component_model::{
    ComponentDescriptor, ComponentItemKind, ConstructorTarget, MethodBody,
};
use vybe_runtime::component_model::{
    InstanceMethodTarget, InstancePropertyTarget, StaticPropertyTarget,
};
pub use winforms::classes;
pub use winforms::dotnet_winforms_component_descriptor;
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StaticMethodTarget {
    Host { module: String, func: String },
    Common { emit: String },
}
pub struct DotnetSurface {
    default_imports: Vec<String>,
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

/// Normalize a source-language receiver type to its descriptor base name.
/// Returns `(base, is_array)`:
/// - `"List<int>"` / `"List(Of Integer)"` → `("List", false)` (generic args stripped)
/// - `"int[]"` / `"Integer()"` → `(original, true)` (array shape)
/// - `"IEnumerable<T>"` → `("IEnumerable", false)`
fn normalize_receiver_type_name(name: &str) -> (String, bool) {
    let trimmed = name.trim();
    // C# `T[]` / `T[,]` and VB `T()` array declarations.
    if trimmed.ends_with(']') && trimmed.contains('[') {
        return (trimmed.to_string(), true);
    }
    if trimmed.ends_with("()") {
        return (trimmed.to_string(), true);
    }
    // Strip generic arguments: `List<int>` / `List(Of Integer)` → `List`.
    let angle = trimmed.find('<');
    let vb = trimmed.to_ascii_lowercase().find("(of");
    let end = match (angle, vb) {
        (Some(a), Some(b)) => a.min(b),
        (Some(a), None) => a,
        (None, Some(b)) => b,
        (None, None) => trimmed.len(),
    };
    let base = trimmed[..end].trim();
    (base.to_string(), false)
}

/// True when the (already generic-stripped) type name denotes something that
/// implements `IEnumerable<T>`, so LINQ resolves against the shared surface.
fn is_enumerable_type_name(name: &str) -> bool {
    let short = name.rsplit('.').next().unwrap_or(name);
    matches!(
        short.to_ascii_lowercase().as_str(),
        "ienumerable"
            | "icollection"
            | "ilist"
            | "ireadonlylist"
            | "ireadonlycollection"
            | "list"
            | "arraylist"
            | "array"
            | "hashset"
            | "sortedset"
            | "queue"
            | "stack"
            | "linkedlist"
            | "concurrentqueue"
            | "concurrentstack"
            | "concurrentbag"
            | "blockingcollection"
            | "observablecollection"
            | "readonlyobservablecollection"
    )
}

fn is_linq_method_name(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "select"
            | "selectmany"
            | "where"
            | "count"
            | "sum"
            | "any"
            | "all"
            | "contains"
            | "reverse"
            | "skip"
            | "skipwhile"
            | "skiplast"
            | "take"
            | "takewhile"
            | "takelast"
            | "first"
            | "firstordefault"
            | "last"
            | "lastordefault"
            | "single"
            | "singleordefault"
            | "elementat"
            | "elementatordefault"
            | "orderby"
            | "orderbydescending"
            | "thenby"
            | "thenbydescending"
            | "distinct"
            | "distinctby"
            | "union"
            | "intersect"
            | "except"
            | "concat"
            | "zip"
            | "toarray"
            | "tolist"
            | "todictionary"
            | "tolookup"
            | "cast"
            | "oftype"
            | "asenumerable"
            | "defaultifempty"
            | "groupby"
            | "min"
            | "max"
            | "average"
            | "aggregate"
    )
}

static DOTNET_SURFACE_CACHE: LazyLock<DotnetSurface> = LazyLock::new(build_dotnet_surface);

fn build_dotnet_surface() -> DotnetSurface {
    let component_descriptor = dotnet_component_descriptor();
    DotnetSurface {
        default_imports: imports::default_interface_imports(),
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
        let (base, _) = normalize_receiver_type_name(name);
        let short = base.rsplit('.').next().unwrap_or(&base);
        self.component_descriptor
            .classes
            .iter()
            .find(|class| {
                class.name.eq_ignore_ascii_case(&base) || class.name.eq_ignore_ascii_case(short)
            })
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
        // Normalize the receiver type: strip generic args (`List<int>` →
        // `List`) so descriptor lookup matches, and detect array/enumerable
        // shapes (`int[]`, `Integer()`, `IEnumerable<T>`).
        let (base, is_array) = normalize_receiver_type_name(class_name);
        let base_short = base.rsplit('.').next().unwrap_or(&base);

        // 1. Method defined directly on the receiver's own class
        //    (`Dictionary.Add`, `Stack.Push`, …). Skip for bare arrays.
        if !is_array {
            if let Some(target) = self.find_instance_method_on(&base, method_name, arg_count) {
                return Some(target);
            }
            if base_short != base {
                if let Some(target) =
                    self.find_instance_method_on(base_short, method_name, arg_count)
                {
                    return Some(target);
                }
            }
        }

        // 2. `System.Linq.Enumerable` fallback — every enumerable receiver
        //    (array, `List<T>`, `HashSet<T>`, `Queue<T>`, query result, …)
        //    resolves LINQ against the single shared `IEnumerable` surface.
        if is_array || is_enumerable_type_name(base_short) {
            return self.find_instance_method_on("IEnumerable", method_name, arg_count);
        }
        if is_linq_method_name(method_name) {
            return self.find_instance_method_on("IEnumerable", method_name, arg_count);
        }
        None
    }

    /// True if `class_name` names a class in the .NET component descriptor
    /// (a framework type like `Button`/`Control`), as opposed to a user class.
    pub fn is_descriptor_class(&self, class_name: &str) -> bool {
        let short = class_name.rsplit('.').next().unwrap_or(class_name);
        self.component_descriptor.classes.iter().any(|class| {
            class.name.eq_ignore_ascii_case(class_name) || class.name.eq_ignore_ascii_case(short)
        })
    }

    /// Find an instance method by name + arity on the named descriptor class or
    /// any of its ancestors (`Button.Show` resolves `Show` on `Control`).
    fn find_instance_method_on(
        &self,
        class_name: &str,
        method_name: &str,
        arg_count: u8,
    ) -> Option<InstanceMethodTarget> {
        let mut current = self
            .component_descriptor
            .classes
            .iter()
            .find(|class| class.name.eq_ignore_ascii_case(class_name));
        while let Some(class) = current {
            // Overload resolution is by exact arity — the adapter class declares
            // each overload (`Count()` runtime property vs `Count(pred)` LINQ).
            // A missing arity must NOT fall back to a different overload.
            if let Some(method) = class.methods.iter().find(|method| {
                !method.is_static
                    && method.name.eq_ignore_ascii_case(method_name)
                    && method.arity == arg_count
            }) {
                return match &method.body {
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
                };
            }
            current = class.parent.as_deref().and_then(|parent| {
                self.component_descriptor
                    .classes
                    .iter()
                    .find(|candidate| candidate.name.eq_ignore_ascii_case(parent))
            });
        }
        None
    }

    pub fn lookup_instance_method_return_type(
        &self,
        class_name: &str,
        method_name: &str,
        arg_count: u8,
    ) -> Option<String> {
        // Mirror `lookup_instance_method`'s normalization so a `var` holding an
        // array/`List<T>`/enumerable (`int[]`, `int()`, `List<int>`) resolves
        // LINQ return types against the shared `IEnumerable` surface — this is
        // what lets a `var` bound to `xs.Skip(2)` chain into `.First()`.
        let (base, is_array) = normalize_receiver_type_name(class_name);
        let base_short = base.rsplit('.').next().unwrap_or(&base);

        if !is_array {
            if let Some(rt) = self.find_return_type_on(&base, method_name, arg_count) {
                return Some(rt);
            }
            if base_short != base {
                if let Some(rt) = self.find_return_type_on(base_short, method_name, arg_count) {
                    return Some(rt);
                }
            }
        }
        if is_array || is_enumerable_type_name(base_short) {
            return self.find_return_type_on("IEnumerable", method_name, arg_count);
        }
        None
    }

    fn find_return_type_on(
        &self,
        class_name: &str,
        method_name: &str,
        arg_count: u8,
    ) -> Option<String> {
        self.component_descriptor
            .classes
            .iter()
            .find(|class| class.name.eq_ignore_ascii_case(class_name))
            .and_then(|class| {
                class
                    .methods
                    .iter()
                    // Exact-arity overload resolution (see `find_instance_method_on`).
                    .find(|method| {
                        !method.is_static
                            && method.name.eq_ignore_ascii_case(method_name)
                            && method.arity == arg_count
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
        self.lookup_instance_accessor(class_name, property_name, false)
    }

    pub fn lookup_instance_property_setter(
        &self,
        class_name: &str,
        property_name: &str,
    ) -> Option<InstancePropertyTarget> {
        self.lookup_instance_accessor(class_name, property_name, true)
    }

    /// Resolve a property accessor by walking `class_name` and its ancestors —
    /// `Button.Text` finds `Text` on `Control`. A generic property accessor
    /// takes the PascalCase key as an argument, so it rides along in `key`;
    /// dedicated per-property host fns leave `key` `None`.
    fn lookup_instance_accessor(
        &self,
        class_name: &str,
        property_name: &str,
        want_setter: bool,
    ) -> Option<InstancePropertyTarget> {
        let requested = class_name.trim();
        let requested_short = requested.rsplit('.').next().unwrap_or(requested);
        if requested_short.eq_ignore_ascii_case("StringBuilder") {
            let emit = match (want_setter, property_name.to_ascii_lowercase().as_str()) {
                (false, "length") => Some("dotnet.sb_length"),
                (true, "length") => Some("dotnet.sb_set_length"),
                (false, "capacity") => Some("dotnet.sb_capacity"),
                (true, "capacity") => Some("dotnet.sb_set_capacity"),
                (false, "maxcapacity") => Some("dotnet.sb_max_capacity"),
                _ => None,
            };
            if let Some(emit) = emit {
                return Some(InstancePropertyTarget::Common {
                    emit: emit.to_string(),
                });
            }
        }
        if requested_short.eq_ignore_ascii_case("Stopwatch") && !want_setter {
            let emit = match property_name.to_ascii_lowercase().as_str() {
                "elapsedmilliseconds" => Some("dotnet.stopwatch_elapsed_ms"),
                "elapsedticks" => Some("dotnet.stopwatch_elapsed_ticks"),
                "elapsed" => Some("dotnet.stopwatch_elapsed"),
                "isrunning" => Some("dotnet.stopwatch_is_running"),
                _ => None,
            };
            if let Some(emit) = emit {
                return Some(InstancePropertyTarget::Common {
                    emit: emit.to_string(),
                });
            }
        }
        if requested_short.eq_ignore_ascii_case("Task") && !want_setter {
            let emit = match property_name.to_ascii_lowercase().as_str() {
                "result" => Some("dotnet.task_result"),
                "iscompleted" => Some("dotnet.task_is_completed"),
                "iscanceled" => Some("dotnet.task_is_canceled"),
                _ => None,
            };
            if let Some(emit) = emit {
                return Some(InstancePropertyTarget::Common {
                    emit: emit.to_string(),
                });
            }
        }
        if requested_short.eq_ignore_ascii_case("DateTime") && !want_setter {
            let emit = match property_name.to_ascii_lowercase().as_str() {
                "year" => Some("dotnet.datetime_year"),
                "month" => Some("dotnet.datetime_month"),
                "day" => Some("dotnet.datetime_day"),
                "hour" => Some("dotnet.datetime_hour"),
                "minute" => Some("dotnet.datetime_minute"),
                "second" => Some("dotnet.datetime_second"),
                "millisecond" => Some("dotnet.datetime_millisecond"),
                "dayofyear" => Some("dotnet.datetime_day_of_year"),
                "dayofweek" => Some("dotnet.datetime_day_of_week"),
                "ticks" => Some("dotnet.datetime_ticks"),
                "kind" => Some("dotnet.datetime_kind"),
                "date" => Some("dotnet.datetime_date"),
                "timeofday" => Some("dotnet.datetime_time_of_day"),
                _ => None,
            };
            if let Some(emit) = emit {
                return Some(InstancePropertyTarget::Common {
                    emit: emit.to_string(),
                });
            }
        }
        if matches!(
            requested_short.to_ascii_lowercase().as_str(),
            "list" | "arraylist"
        ) && !want_setter
            && property_name.eq_ignore_ascii_case("Capacity")
        {
            return Some(InstancePropertyTarget::Common {
                emit: "dotnet.list_capacity".to_string(),
            });
        }
        let mut current = self.component_descriptor.classes.iter().find(|class| {
            class.name.eq_ignore_ascii_case(requested)
                || class.name.eq_ignore_ascii_case(requested_short)
        });
        while let Some(class) = current {
            if let Some(property) = class
                .properties
                .iter()
                .find(|property| property.name.eq_ignore_ascii_case(property_name))
            {
                let target = if want_setter {
                    property.setter.as_ref()
                } else {
                    property.getter.as_ref()
                };
                return target.map(|target| {
                    let keyed = target.name == vybe_compiler::primitives::gui::HOST_FN_GET_PROPERTY
                        || target.name == vybe_compiler::primitives::gui::HOST_FN_SET_PROPERTY;
                    InstancePropertyTarget::Host {
                        module: target.module.clone(),
                        func: target.name.clone(),
                        key: keyed.then(|| property.name.clone()),
                    }
                });
            }
            current = class.parent.as_deref().and_then(|parent| {
                self.component_descriptor
                    .classes
                    .iter()
                    .find(|candidate| candidate.name.eq_ignore_ascii_case(parent))
            });
        }
        None
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

/// The declared return type of an INSTANCE member, for tree registration.
/// Public so `tree_register` can declare it with the class.
pub fn instance_method_return_type(class_name: &str, method_name: &str) -> Option<String> {
    dotnet_instance_method_return_type(class_name, method_name)
}

fn dotnet_instance_method_return_type(class_name: &str, method_name: &str) -> Option<String> {
    let class = class_name.rsplit('.').next().unwrap_or(class_name);
    if class.eq_ignore_ascii_case("TimeZoneInfo") {
        return match method_name.to_ascii_lowercase().as_str() {
            "getutcoffset" => Some("TimeSpan".into()),
            "isdaylightsavingtime" => Some("Boolean".into()),
            _ => None,
        };
    }
    // ── Value types and builders ──────────────────────────────────────────
    //
    // A leaf that does not say what it RETURNS ends the chain: the value it
    // produces is untyped, so the next member resolves against nothing. These
    // are the classes whose methods return their own type (or another declared
    // one), which is what makes `d.AddDays(1).AddHours(2).ToString("s")` and
    // `sb.Append(x).AppendLine(y)` resolve past the first hop.
    if class.eq_ignore_ascii_case("DateTime") {
        return match method_name.to_ascii_lowercase().as_str() {
            "add" | "adddays" | "addhours" | "addminutes" | "addseconds"
            | "addmilliseconds" | "addmonths" | "addyears" | "addticks" | "date"
            | "touniversaltime" | "tolocaltime" | "fromoadate" | "frombinary"
            | "fromfiletimeutc" | "specifykind" | "parse" | "parseexact" | "now" | "utcnow"
            | "today" | "minvalue" | "maxvalue" => Some("DateTime".into()),
            "subtract" | "timeofday" => Some("TimeSpan".into()),
            "tostring" | "toshortdatestring" | "tolongdatestring" | "toshorttimestring"
            | "tolongtimestring" => Some("String".into()),
            "compareto" | "compare" | "year" | "month" | "day" | "hour" | "minute" | "second"
            | "millisecond" | "dayofyear" | "daysinmonth" => Some("Int32".into()),
            "equals" | "isleapyear" | "isdaylightsavingtime" => Some("Boolean".into()),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("TimeSpan") {
        return match method_name.to_ascii_lowercase().as_str() {
            "add" | "subtract" | "negate" | "duration" | "fromdays" | "fromhours"
            | "fromminutes" | "fromseconds" | "frommilliseconds" | "fromticks" | "parse"
            | "minvalue" | "maxvalue" | "multiply" | "divide" => Some("TimeSpan".into()),
            "tostring" => Some("String".into()),
            "compareto" | "compare" | "days" | "hours" | "minutes" | "seconds"
            | "milliseconds" => Some("Int32".into()),
            "equals" => Some("Boolean".into()),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("DateTimeOffset") {
        return match method_name.to_ascii_lowercase().as_str() {
            "add" | "adddays" | "addhours" | "addminutes" | "addseconds"
            | "addmilliseconds" | "addmonths" | "addyears" | "addticks" | "touniversaltime"
            | "tolocaltime" | "parse" | "now" | "utcnow" => Some("DateTimeOffset".into()),
            "subtract" | "offset" => Some("TimeSpan".into()),
            "datetime" | "utcdatetime" | "localdatetime" | "date" => Some("DateTime".into()),
            "tostring" => Some("String".into()),
            "compareto" | "compare" => Some("Int32".into()),
            "equals" => Some("Boolean".into()),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("List") || class.eq_ignore_ascii_case("ArrayList") {
        return match method_name.to_ascii_lowercase().as_str() {
            "getrange" | "findall" | "convertall" | "toarray" | "tolist" | "asreadonly"
            | "distinct" | "where" | "select" | "orderby" | "orderbydescending" | "take"
            | "skip" | "reverse" | "concat" => Some("List".into()),
            "tostring" => Some("String".into()),
            "count" | "capacity" | "indexof" | "lastindexof" | "findindex"
            | "findlastindex" | "binarysearch" | "removeall" => Some("Int32".into()),
            "contains" | "exists" | "trueforall" | "any" | "all" | "remove" => {
                Some("Boolean".into())
            }
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Dictionary") {
        return match method_name.to_ascii_lowercase().as_str() {
            "keys" | "values" | "entries" | "entriessorted" => Some("List".into()),
            "tostring" => Some("String".into()),
            "count" => Some("Int32".into()),
            "containskey" | "containsvalue" | "tryadd" | "remove" | "trygetvalue" => {
                Some("Boolean".into())
            }
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Uri") || class.eq_ignore_ascii_case("UriBuilder") {
        return match method_name.to_ascii_lowercase().as_str() {
            "uri" => Some("Uri".into()),
            "tostring" | "absoluteuri" | "absolutepath" | "host" | "scheme" | "query"
            | "fragment" | "localpath" | "authority" | "pathandquery" | "originalstring"
            | "getleftpart" | "unescapedatastring" | "escapedatastring" => Some("String".into()),
            "port" => Some("Int32".into()),
            "isabsoluteuri" | "isdefaultport" | "isfile" | "isloopback" => Some("Boolean".into()),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Version") {
        return match method_name.to_ascii_lowercase().as_str() {
            "parse" => Some("Version".into()),
            "tostring" => Some("String".into()),
            "major" | "minor" | "build" | "revision" | "compareto" => Some("Int32".into()),
            "equals" => Some("Boolean".into()),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Guid") {
        return match method_name.to_ascii_lowercase().as_str() {
            "newguid" | "parse" | "empty" => Some("Guid".into()),
            "tostring" => Some("String".into()),
            "equals" => Some("Boolean".into()),
            "compareto" => Some("Int32".into()),
            _ => None,
        };
    }
    // ── System.String ─────────────────────────────────────────────────────
    //
    // A leaf is not complete until it declares what it RETURNS. `s.Trim()`
    // resolved fine and `s.Trim().Substring(0, 5)` did not, because the value
    // `Trim` produced carried no type and the next hop had nothing to resolve
    // against — the same failure as the XML and ADO chains below.
    //
    // It only became visible when the `[value_methods]` rows came out: a
    // profile row answers on ANY receiver, so while it shadowed these leaves
    // the missing return type could not be observed.
    if class.eq_ignore_ascii_case("String") {
        return match method_name.to_ascii_lowercase().as_str() {
            "trim" | "trimstart" | "trimend" | "toupper" | "tolower" | "substring" | "replace"
            | "insert" | "remove" | "padleft" | "padright" | "tostring" | "normalize" => {
                Some("String".into())
            }
            // ⛔ NOT `String()`. The array spelling makes the VB walker treat
            // the local as an array and rewrite `.Length` — the same trap the
            // `XElement.Elements()` note below records. Left undeclared until
            // there is a sequence spelling that does not retype the local.
            "indexof" | "lastindexof" | "indexofany" | "lastindexofany" | "compareto"
            | "length" => Some("Int32".into()),
            "contains" | "startswith" | "endswith" | "equals" => Some("Boolean".into()),
            _ => None,
        };
    }
    // ── System.Xml.Linq ───────────────────────────────────────────────────
    //
    // Same failure as the ADO chains below, one layer up: the methods were
    // declared, what they RETURN was not. `Dim c = d.Elements()` bound a value
    // with no type, so `c.First()` could not reach the shared `IEnumerable`
    // surface and died as `RuntimeError: null` — while `c.Length` answered 2,
    // because a raw property read needs no type.
    if class.eq_ignore_ascii_case("XElement") || class.eq_ignore_ascii_case("XDocument") {
        return match method_name.to_ascii_lowercase().as_str() {
            // ⛔ `IEnumerable`, NOT `XElement()`. The array spelling also
            // satisfies the dotnet resolver, but it makes the VB walker treat
            // the local as an array and rewrite `.Length` — which then answered
            // 0 for a sequence of 2. Measured both ways.
            "elements" | "descendants" | "ancestors" | "nodes" | "attributes"
            | "elementsafterself" | "elementsbeforeself" => Some("IEnumerable".into()),
            "element" | "root" | "parent" => Some("XElement".into()),
            "attribute" => Some("XAttribute".into()),
            _ => None,
        };
    }
    // ── ADO.NET ───────────────────────────────────────────────────────────
    //
    // Every ADO chain is a chain of FACTORY calls — `conn.CreateCommand()`,
    // `cmd.ExecuteReader()` — and each hop was undeclared, so the value came
    // back with no type and the next member on it resolved against nothing.
    // `Dim cmd As Object = conn.CreateCommand()` is the shape every ADO
    // example uses, and `cmd.ExecuteNonQuery()` reached
    // `undefined is not callable`.
    //
    // The methods themselves have always been declared (they are `wasi:sql`
    // host calls in `component_classes_data_drawing.rs`); what was missing is
    // what they RETURN.
    if class.eq_ignore_ascii_case("SqlConnection")
        || class.eq_ignore_ascii_case("OleDbConnection")
    {
        return match method_name.to_ascii_lowercase().as_str() {
            "createcommand" => Some("SqlCommand".into()),
            "begintransaction" => Some("SqlTransaction".into()),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("SqlCommand") || class.eq_ignore_ascii_case("OleDbCommand") {
        // The `*Async` twins return the same thing here: the adapter awaits at
        // the boundary, so the value a program binds is the reader/count, not
        // a Task. Declaring only the sync half left `ExecuteReaderAsync()`
        // untyped and `rdr.Read()` unresolvable — the same failure one call
        // further along.
        return match method_name.to_ascii_lowercase().as_str() {
            "executereader" | "executereaderasync" => Some("SqlDataReader".into()),
            "executenonquery" | "executenonqueryasync" => Some("Int32".into()),
            "createparameter" => Some("SqlParameter".into()),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("SqlDataReader")
        || class.eq_ignore_ascii_case("OleDbDataReader")
    {
        return match method_name.to_ascii_lowercase().as_str() {
            "read" | "nextresult" | "isdbnull" => Some("Boolean".into()),
            "getstring" | "getname" => Some("string".into()),
            "getint32" | "getfieldcount" | "fieldcount" => Some("Int32".into()),
            "getdouble" => Some("Double".into()),
            _ => None,
        };
    }
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
    if class.eq_ignore_ascii_case("StreamReader") {
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "readline" | "readtoend"
        ) {
            return Some("string".into());
        }
        if method_name.eq_ignore_ascii_case("EndOfStream") {
            return Some("Boolean".into());
        }
    }
    if class.eq_ignore_ascii_case("FileStream") && method_name.eq_ignore_ascii_case("Read") {
        return Some("string".into());
    }
    if class.eq_ignore_ascii_case("StringWriter") && method_name.eq_ignore_ascii_case("ToString") {
        return Some("string".into());
    }
    if class.eq_ignore_ascii_case("StringWriter")
        && method_name.eq_ignore_ascii_case("GetStringBuilder")
    {
        return Some("StringBuilder".into());
    }
    if class.eq_ignore_ascii_case("StringWriter")
        && method_name.eq_ignore_ascii_case("WriteLineAsync")
    {
        return Some("Task".into());
    }
    if class.eq_ignore_ascii_case("DateTime") {
        return match method_name.to_ascii_lowercase().as_str() {
            "add" | "adddays" | "addhours" | "addmonths" | "addyears" | "touniversaltime"
            | "tolocaltime" => Some("DateTime".into()),
            "subtract" => Some("TimeSpan".into()),
            "compareto" | "gethashcode" => Some("Int32".into()),
            "equals" => Some("Boolean".into()),
            "tostring" | "toshortdatestring" => Some("string".into()),
            "tobinary" | "tofiletimeutc" => Some("Int64".into()),
            "tooadate" => Some("Double".into()),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("DateTimeOffset") {
        return match method_name.to_ascii_lowercase().as_str() {
            "adddays" | "addhours" | "tooffset" | "touniversaltime" => {
                Some("DateTimeOffset".into())
            }
            "compareto" | "gethashcode" => Some("Int32".into()),
            "equals" | "equalsexact" => Some("Boolean".into()),
            "tostring" => Some("string".into()),
            "tounixtimeseconds" | "tounixtimemilliseconds" => Some("Int64".into()),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("XElement") {
        if matches!(method_name.to_ascii_lowercase().as_str(), "element") {
            return Some("XElement".into());
        }
        if matches!(method_name.to_ascii_lowercase().as_str(), "elements") {
            return Some("IEnumerable".into());
        }
        if matches!(method_name.to_ascii_lowercase().as_str(), "name") {
            return Some("XName".into());
        }
        if method_name.eq_ignore_ascii_case("attribute") {
            return Some("XAttribute".into());
        }
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "value" | "tostring"
        ) {
            return Some("string".into());
        }
    }
    if class.eq_ignore_ascii_case("XDocument") {
        if matches!(method_name.to_ascii_lowercase().as_str(), "root") {
            return Some("XElement".into());
        }
    }
    if class.eq_ignore_ascii_case("XName") {
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "localname" | "namespacename" | "tostring"
        ) {
            return Some("string".into());
        }
    }
    if class.eq_ignore_ascii_case("CancellationTokenSource") {
        if method_name.eq_ignore_ascii_case("Token") {
            return Some("CancellationToken".into());
        }
    }
    if class.eq_ignore_ascii_case("CancellationToken") {
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "register" | "waithandle"
        ) {
            return Some("Object".into());
        }
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "iscancellationrequested" | "canbecanceled"
        ) {
            return Some("Boolean".into());
        }
    }
    if class.eq_ignore_ascii_case("Task") {
        if method_name.eq_ignore_ascii_case("iscanceled") {
            return Some("Boolean".into());
        }
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "wait" | "continuewith"
        ) {
            return Some("Task".into());
        }
    }
    if class.eq_ignore_ascii_case("ValueTask") && method_name.eq_ignore_ascii_case("AsTask") {
        return Some("Task".into());
    }
    if class.eq_ignore_ascii_case("Process") && method_name.eq_ignore_ascii_case("WaitForExit") {
        return Some("Boolean".into());
    }
    // LINQ deferred (sequence-returning) operators stay `IEnumerable<T>`, so a
    // chain like `xs.OrderBy(k).Distinct().Where(p)` keeps resolving each step
    // against the shared surface. Terminal operators (`Count`, `Sum`, `First`,
    // `ToList`, …) are intentionally excluded — they return scalars/collections.
    // Normalize the array shape first (`int[]` / `int()` → array) so a `var`
    // holding an array literal chains as well as an explicit `IEnumerable`.
    let (base, is_array) = normalize_receiver_type_name(class_name);
    let base_short = base.rsplit('.').next().unwrap_or(&base);
    if is_array || is_enumerable_type_name(base_short) {
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "all" | "any" | "contains" | "sequenceequal"
        ) {
            return Some("Boolean".into());
        }
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "count" | "longcount"
        ) {
            return Some("Int32".into());
        }
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "sum" | "average" | "min" | "max"
        ) {
            return Some("Double".into());
        }
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "first"
                | "firstordefault"
                | "last"
                | "lastordefault"
                | "single"
                | "singleordefault"
                | "elementat"
                | "elementatordefault"
        ) {
            return Some("Object".into());
        }
        if matches!(method_name.to_ascii_lowercase().as_str(), "tolist") {
            return Some("List".into());
        }
        if matches!(method_name.to_ascii_lowercase().as_str(), "toarray") {
            return Some("Array".into());
        }
        if matches!(
            method_name.to_ascii_lowercase().as_str(),
            "where"
                | "select"
                | "selectmany"
                | "distinct"
                | "distinctby"
                | "orderby"
                | "orderbydescending"
                | "thenby"
                | "thenbydescending"
                | "skip"
                | "skipwhile"
                | "skiplast"
                | "take"
                | "takewhile"
                | "takelast"
                | "reverse"
                | "concat"
                | "union"
                | "intersect"
                | "except"
                | "append"
                | "prepend"
                | "defaultifempty"
                | "groupby"
                | "zip"
                | "cast"
                | "oftype"
                | "asenumerable"
        ) {
            return Some("IEnumerable".into());
        }
    }
    None
}

pub fn default_interface_imports() -> Vec<String> {
    surface().default_imports().to_vec()
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
    static CONSTANTS: LazyLock<Vec<(&'static str, f64)>> = LazyLock::new(|| {
        let mut constants = Vec::new();
        constants.extend_from_slice(core::types::namespace_constants());
        constants.extend_from_slice(winforms::types::namespace_constants());
        constants
    });
    CONSTANTS.as_slice()
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

pub fn is_component_descriptor_class(name: &str) -> bool {
    surface().is_descriptor_class(name)
}

pub fn component_descriptor_class_interface(name: &str) -> Option<String> {
    let short = name.rsplit('.').next().unwrap_or(name);
    surface()
        .component_descriptor
        .exports
        .iter()
        .find_map(|export| {
            let ComponentItemKind::Class(class) = &export.kind else {
                return None;
            };
            if class.name.eq_ignore_ascii_case(name) || class.name.eq_ignore_ascii_case(short) {
                Some(export.interface.clone())
            } else {
                None
            }
        })
}

pub fn is_component_descriptor_class_in_namespace(name: &str, namespace_prefix: &str) -> bool {
    let short = name.rsplit('.').next().unwrap_or(name);
    surface().component_descriptor.exports.iter().any(|export| {
        if !export.interface.starts_with(namespace_prefix) {
            return false;
        }
        let ComponentItemKind::Class(class) = &export.kind else {
            return false;
        };
        class.name.eq_ignore_ascii_case(name) || class.name.eq_ignore_ascii_case(short)
    })
}

pub fn component_instance_method_exists(
    class_name: &str,
    method_name: &str,
    arg_count: u8,
) -> bool {
    lookup_component_instance_method(class_name, method_name, arg_count).is_some()
}

/// The bare type name at the end of a .NET type path, with any generic
/// argument list dropped — `System.Collections.Generic.Comparer<int>` and
/// `Comparer(Of Integer)` both yield `Comparer`.
///
/// Every .NET frontend spells type arguments differently; the type they name
/// is the same, so a table keyed on the type must not see the spelling.
fn dotnet_type_path_leaf(path: &str) -> String {
    let path = path.trim();
    let base = match (path.find('<'), path.find("(Of ")) {
        (Some(i), Some(j)) => &path[..i.min(j)],
        (Some(i), None) => &path[..i],
        (None, Some(j)) => &path[..j],
        (None, None) => path,
    };
    base.trim()
        .rsplit('.')
        .next()
        .unwrap_or(base)
        .trim()
        .to_string()
}

/// Does a generic call on this .NET type need its type ARGUMENTS materialized
/// as extra runtime arguments?
///
/// A C# frontend appends two positional bindings per type argument (a ctor
/// handle and a zero value) so a generic method can evaluate `new T()` or
/// `default(T)` at run time. A USER generic needs that. A declared .NET static
/// never does — its body is a leaf in this platform, which reads its operands
/// positionally, so the bindings land as DATA:
/// `ImmutableArray.Create<int>(1)` arrived with argc 3 and built a
/// three-element array whose `Length` answered 3.
///
/// Answered from the descriptors rather than a name list, so a type gains the
/// right behaviour by being declared — the check is "does this platform
/// declare a static of that name on that type", which is exactly the
/// condition under which the leaf, not the caller, owns the arity.
pub fn generic_binding_args_suppressed(class_path: &str, method_name: &str) -> bool {
    let class = dotnet_type_path_leaf(class_path);
    if class.is_empty() {
        return false;
    }
    class_exports::dotnet_class_exports()
        .iter()
        .filter(|export| export.class.name.eq_ignore_ascii_case(&class))
        .flat_map(|export| export.class.methods.iter())
        .any(|m| m.is_static && m.name.eq_ignore_ascii_case(method_name))
}

/// Is this one of the built-in comparer singletons?
///
/// `Comparer.Default`, `EqualityComparer.Default`, `StringComparer.Ordinal`
/// and `StringComparer.OrdinalIgnoreCase` — the reads that used to answer a
/// raw marker string. Keyed on the type leaf so a closed generic
/// (`Comparer<int>`, `Comparer(Of Integer)`) answers the same as the open one.
fn is_comparer_static(class_path: &str, member_name: &str) -> bool {
    match dotnet_type_path_leaf(class_path).to_ascii_lowercase().as_str() {
        "comparer" | "equalitycomparer" => member_name.eq_ignore_ascii_case("Default"),
        "stringcomparer" => {
            member_name.eq_ignore_ascii_case("Ordinal")
                || member_name.eq_ignore_ascii_case("OrdinalIgnoreCase")
        }
        _ => false,
    }
}

/// What an `Immutable*` static factory hands back.
///
/// `CreateBuilder` is the exception that proves the rule: it is the only one
/// that does NOT answer its own type, because a builder is the mutable staging
/// object you call `ToImmutable()` on.
fn immutable_factory_return_type(class: &str, method_name: &str) -> Option<&'static str> {
    let owner = match class.to_ascii_lowercase().as_str() {
        "immutablearray" => "ImmutableArray",
        "immutablelist" => "ImmutableList",
        "immutablequeue" => "ImmutableQueue",
        "immutablestack" => "ImmutableStack",
        "immutablehashset" => "ImmutableHashSet",
        "immutabledictionary" => "ImmutableDictionary",
        _ => return None,
    };
    match method_name.to_ascii_lowercase().as_str() {
        "createbuilder" => Some("ImmutableListBuilder"),
        "create" | "createrange" | "empty" => Some(owner),
        _ => None,
    }
}

pub fn static_member_constant(prefix: &str, member_name: &str) -> Option<&'static str> {
    let normalized = prefix.trim();
    if (normalized.eq_ignore_ascii_case("StringComparison")
        || normalized.eq_ignore_ascii_case("System.StringComparison"))
        && member_name.eq_ignore_ascii_case("OrdinalIgnoreCase")
    {
        return Some("__dotnet_stringcomparison_ordinalignorecase");
    }
    if (normalized.eq_ignore_ascii_case("StringComparison")
        || normalized.eq_ignore_ascii_case("System.StringComparison"))
        && member_name.eq_ignore_ascii_case("InvariantCultureIgnoreCase")
    {
        return Some("__dotnet_stringcomparison_invariantignorecase");
    }
    if (normalized.eq_ignore_ascii_case("StringComparison")
        || normalized.eq_ignore_ascii_case("System.StringComparison"))
        && (member_name.eq_ignore_ascii_case("Ordinal")
            || member_name.eq_ignore_ascii_case("InvariantCulture")
            || member_name.eq_ignore_ascii_case("CurrentCulture"))
    {
        return Some("__dotnet_stringcomparison_ordinal");
    }
    if normalized.eq_ignore_ascii_case("Base64FormattingOptions")
        || normalized.eq_ignore_ascii_case("System.Base64FormattingOptions")
    {
        if member_name.eq_ignore_ascii_case("InsertLineBreaks") {
            return Some("__dotnet_base64_insertlinebreaks");
        }
        if member_name.eq_ignore_ascii_case("None") {
            return Some("__dotnet_base64_none");
        }
    }
    if normalized.eq_ignore_ascii_case("StringSplitOptions")
        || normalized.eq_ignore_ascii_case("System.StringSplitOptions")
    {
        if member_name.eq_ignore_ascii_case("RemoveEmptyEntries") {
            return Some("__dotnet_stringsplit_removeemptyentries");
        }
        if member_name.eq_ignore_ascii_case("None") {
            return Some("__dotnet_stringsplit_none");
        }
    }
    // `MidpointRounding.*` — the member's own name IS the value, same shape as
    // `TypeCode` below.
    //
    // Nothing in the tree defined `MidpointRounding` at all, so every explicit
    // overload — `Math.Round(x, d, MidpointRounding.ToEven)`,
    // `Decimal.Round(x, d, MidpointRounding.AwayFromZero)` — failed to resolve.
    // The POLICIES themselves have been available the whole time as
    // `common:math.round_half_even` / `common:math.round_half_away`; only the
    // name that SELECTS one was missing. Kept as the .NET spelling so the
    // selecting lowering can match on it without a second representation.
    if normalized.eq_ignore_ascii_case("MidpointRounding")
        || normalized.eq_ignore_ascii_case("System.MidpointRounding")
    {
        return Some(match member_name.to_ascii_lowercase().as_str() {
            "toeven" => "ToEven",
            "awayfromzero" => "AwayFromZero",
            _ => return None,
        });
    }
    // `Microsoft.VisualBasic.DateInterval.*` — the member's own name IS the
    // value, same shape as `TypeCode` below.
    //
    // `DateAdd(DateInterval.Day, …)` and `DateAdd("d", …)` are the same call
    // in VB, so the enum member has to arrive as an ordinary value that the
    // DateAndTime adapter's interval tables can match — which they do, case
    // insensitively, against both this spelling and the abbreviation.
    if normalized.eq_ignore_ascii_case("DateInterval")
        || normalized.eq_ignore_ascii_case("Microsoft.VisualBasic.DateInterval")
    {
        return Some(match member_name.to_ascii_lowercase().as_str() {
            "year" => "Year",
            "quarter" => "Quarter",
            "month" => "Month",
            "dayofyear" => "DayOfYear",
            "day" => "Day",
            "weekofyear" => "WeekOfYear",
            "weekday" => "Weekday",
            "hour" => "Hour",
            "minute" => "Minute",
            "second" => "Second",
            _ => return None,
        });
    }
    // `TypeCode.*` — the member's own name IS the value.
    //
    // `Type.GetTypeCode(t)` (resolved in `primitives/reflection.rs` off the
    // .NET type-name leaf) yields the same spelling, so `= TypeCode.Int32`
    // compares equal and `.ToString()` renders it, with no separate enum
    // representation to keep in sync. Nothing in the tree defined `TypeCode`
    // before this — VB's walker answered `Type.GetTypeCode` itself, in VB's
    // own type spelling.
    if normalized.eq_ignore_ascii_case("TypeCode")
        || normalized.eq_ignore_ascii_case("System.TypeCode")
    {
        return Some(match member_name.to_ascii_lowercase().as_str() {
            "empty" => "Empty",
            "object" => "Object",
            "dbnull" => "DBNull",
            "boolean" => "Boolean",
            "char" => "Char",
            "sbyte" => "SByte",
            "byte" => "Byte",
            "int16" => "Int16",
            "uint16" => "UInt16",
            "int32" => "Int32",
            "uint32" => "UInt32",
            "int64" => "Int64",
            "uint64" => "UInt64",
            "single" => "Single",
            "double" => "Double",
            "decimal" => "Decimal",
            "datetime" => "DateTime",
            "string" => "String",
            _ => return None,
        });
    }
    // ⛔ These four markers are NOT merely the value of a comparer read. Both
    // .NET frontends match them as a CONSTRUCTOR ARGUMENT
    // (`new Dictionary<string,int>(StringComparer.OrdinalIgnoreCase)`) to
    // select a case-folding container — see `has_ordinal_ignore_case_comparer_arg`
    // in the C# walker and the HashSet/Dictionary ctor arms in the VB one.
    // No method is called on the comparer there, so a typed tree member with
    // `Equals`/`Compare` does NOT subsume this. Dropping the marker made those
    // containers silently case-SENSITIVE — a wrong answer, not a failure.
    // Retiring it needs comparer-aware Dictionary/HashSet CONSTRUCTORS in this
    // platform first; the ctor currently backs onto `dotnet.dict_new_ignore_arg`,
    // which discards the argument outright.
    if (normalized.eq_ignore_ascii_case("StringComparer")
        || normalized.eq_ignore_ascii_case("System.StringComparer"))
        && member_name.eq_ignore_ascii_case("OrdinalIgnoreCase")
    {
        return Some("__dotnet_stringcomparer_ordinalignorecase");
    }
    if (normalized.eq_ignore_ascii_case("StringComparer")
        || normalized.eq_ignore_ascii_case("System.StringComparer"))
        && member_name.eq_ignore_ascii_case("Ordinal")
    {
        return Some("__dotnet_stringcomparer_ordinal");
    }
    if member_name.eq_ignore_ascii_case("Default") {
        // Keyed on the TYPE, not the closed spelling: the old
        // `normalized.contains("Comparer<")` could only match C#'s generic
        // syntax, so `Comparer(Of Integer).Default` never resolved for VB.
        match dotnet_type_path_leaf(normalized).as_str() {
            leaf if leaf.eq_ignore_ascii_case("EqualityComparer") => {
                return Some("__dotnet_equalitycomparer_default");
            }
            leaf if leaf.eq_ignore_ascii_case("Comparer") => {
                return Some("__dotnet_comparer_default");
            }
            _ => {}
        }
    }
    if (normalized.eq_ignore_ascii_case("DateTimeKind")
        || normalized.eq_ignore_ascii_case("System.DateTimeKind"))
        && (member_name.eq_ignore_ascii_case("Utc")
            || member_name.eq_ignore_ascii_case("Local")
            || member_name.eq_ignore_ascii_case("Unspecified"))
    {
        return Some(match member_name.to_ascii_lowercase().as_str() {
            "utc" => "Utc",
            "local" => "Local",
            _ => "Unspecified",
        });
    }
    if normalized.eq_ignore_ascii_case("DateTimeStyles")
        || normalized.eq_ignore_ascii_case("System.Globalization.DateTimeStyles")
    {
        return match member_name.to_ascii_lowercase().as_str() {
            "none" => Some("None"),
            "allowwhitespaces" => Some("AllowWhiteSpaces"),
            "adjusttouniversal" => Some("AdjustToUniversal"),
            "assumeuniversal" => Some("AssumeUniversal"),
            "roundtripkind" => Some("RoundtripKind"),
            _ => None,
        };
    }
    if (normalized.eq_ignore_ascii_case("NotifyCollectionChangedAction")
        || normalized
            .eq_ignore_ascii_case("System.Collections.Specialized.NotifyCollectionChangedAction"))
        && (member_name.eq_ignore_ascii_case("Add")
            || member_name.eq_ignore_ascii_case("Remove")
            || member_name.eq_ignore_ascii_case("Replace")
            || member_name.eq_ignore_ascii_case("Move")
            || member_name.eq_ignore_ascii_case("Reset"))
    {
        return Some(match member_name.to_ascii_lowercase().as_str() {
            "add" => "Add",
            "remove" => "Remove",
            "replace" => "Replace",
            "move" => "Move",
            _ => "Reset",
        });
    }
    None
}

pub fn static_member_parameterless_call(prefix: &str, member_name: &str) -> bool {
    let trimmed = prefix.trim();
    let normalized_storage;
    let normalized = if trimmed.contains('<') {
        normalized_storage = strip_generic_suffixes(trimmed);
        normalized_storage.as_str()
    } else {
        trimmed
    };
    // `ImmutableDictionary<K,V>.Empty` — no parentheses, so the walker reads it
    // as a FIELD unless told otherwise, and a field that no type declares is
    // `undefined`. Declaring it here makes the read INVOKE the zero-arg static,
    // which is how every other value-shaped .NET static (`DateTime.Now`,
    // `TimeSpan.Zero`) already resolves.
    //
    // `Create`/`CreateRange`/`CreateBuilder` are deliberately absent: they are
    // always written WITH an argument list, so they arrive as calls already.
    if member_name.eq_ignore_ascii_case("Empty")
        && immutable_factory_return_type(&dotnet_type_path_leaf(normalized), member_name).is_some()
    {
        return true;
    }
    if (normalized.eq_ignore_ascii_case("CancellationToken")
        || normalized.eq_ignore_ascii_case("System.Threading.CancellationToken"))
        && member_name.eq_ignore_ascii_case("None")
    {
        return true;
    }
    ((normalized.eq_ignore_ascii_case("DateTime")
        || normalized.eq_ignore_ascii_case("System.DateTime"))
        && (member_name.eq_ignore_ascii_case("Now")
            || member_name.eq_ignore_ascii_case("UtcNow")
            || member_name.eq_ignore_ascii_case("Today")))
        || ((normalized.eq_ignore_ascii_case("TimeSpan")
            || normalized.eq_ignore_ascii_case("System.TimeSpan"))
            && member_name.eq_ignore_ascii_case("Zero"))
        // ⛔ `BigInteger.Zero` is a .NET static PROPERTY, but its value is an
        // object this platform MINTS (payload plus operator slots), so the tree
        // backs it with a zero-arity static METHOD — `class.properties` only
        // accepts a host getter. Without this the member read never invoked the
        // leaf: `BigInteger.Zero.IsZero` answered `undefined` while
        // `.ToString()` still printed `0`, because a bare BigInt stringifies.
        || ((normalized.eq_ignore_ascii_case("BigInteger")
            || normalized.eq_ignore_ascii_case("System.Numerics.BigInteger"))
            && matches!(
                member_name.to_ascii_lowercase().as_str(),
                "zero" | "one" | "minusone"
            ))
        || ((normalized.eq_ignore_ascii_case("Stopwatch")
            || normalized.eq_ignore_ascii_case("System.Diagnostics.Stopwatch"))
            && matches!(
                member_name.to_ascii_lowercase().as_str(),
                "frequency" | "ishighresolution"
            ))
        || ((normalized.eq_ignore_ascii_case("Encoding")
            || normalized.eq_ignore_ascii_case("System.Text.Encoding"))
            && matches!(
                member_name.to_ascii_lowercase().as_str(),
                "utf8" | "ascii" | "unicode" | "default" | "utf32" | "latin1" | "bigendianunicode"
            ))
        || ((normalized.eq_ignore_ascii_case("ArrayPool")
            || normalized.eq_ignore_ascii_case("System.Buffers.ArrayPool"))
            && member_name.eq_ignore_ascii_case("Shared"))
        || ((normalized.eq_ignore_ascii_case("MemoryPool")
            || normalized.eq_ignore_ascii_case("System.Buffers.MemoryPool"))
            && member_name.eq_ignore_ascii_case("Shared"))
}

fn strip_generic_suffixes(name: &str) -> String {
    let mut out = String::with_capacity(name.len());
    let mut depth = 0usize;
    for ch in name.chars() {
        match ch {
            '<' => depth += 1,
            '>' => depth = depth.saturating_sub(1),
            _ if depth == 0 => out.push(ch),
            _ => {}
        }
    }
    out
}

pub fn canonical_type_name(raw: &str) -> String {
    let trimmed = raw.trim().trim_end_matches('?').trim();
    let trimmed = trimmed.strip_suffix("()").unwrap_or(trimmed).trim();
    let base = trimmed
        .split_once("(Of ")
        .map(|(base, _)| base)
        .or_else(|| trimmed.split_once('<').map(|(base, _)| base))
        .unwrap_or(trimmed)
        .trim();
    let leaf = base.rsplit('.').next().unwrap_or(base).trim();
    match leaf.to_ascii_lowercase().as_str() {
        "short" | "int16" => "Int16",
        "integer" | "int" | "int32" => "Int32",
        "long" | "int64" => "Int64",
        "byte" => "Byte",
        "sbyte" => "SByte",
        "ushort" | "uint16" => "UInt16",
        "uinteger" | "uint" | "uint32" => "UInt32",
        "ulong" | "uint64" => "UInt64",
        "single" | "float" => "Single",
        "double" => "Double",
        "decimal" => "Decimal",
        "boolean" | "bool" => "Boolean",
        "string" => "String",
        "char" => "Char",
        "object" => "Object",
        "date" | "datetime" => "DateTime",
        _ => leaf,
    }
    .to_string()
}

pub fn lookup_component_instance_method(
    class_name: &str,
    method_name: &str,
    arg_count: u8,
) -> Option<InstanceMethodTarget> {
    surface().lookup_instance_method(class_name, method_name, arg_count)
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

pub fn static_method_return_type(class_name: &str, method_name: &str) -> Option<&'static str> {
    let class = class_name.rsplit('.').next().unwrap_or(class_name);
    // `Type.GetTypeCode` and `Nullable.GetUnderlyingType` — declared here so the
    // return type comes off the TREE (`member_returns`) rather than a walker's
    // own table. `namespaceplan.md`: "chaining methods declare `member_returns`".
    // The VB walker encoded both, which is `System.*` knowledge living in a
    // language frontend where no other .NET consumer could see it.
    if class.eq_ignore_ascii_case("Type") && method_name.eq_ignore_ascii_case("GetTypeCode") {
        return Some("TypeCode");
    }
    // `System.Collections.Immutable` factories answer their own type — the
    // ENTRY point into a chain that `self_member_returns` then carries. A
    // static returning a bare array has no identity, so `Create(…).Add(…)`
    // resolved against nothing even with both members declared.
    if let Some(immutable) = immutable_factory_return_type(class, method_name) {
        return Some(immutable);
    }
    if is_comparer_static(class, method_name) {
        return Some(match class.to_ascii_lowercase().as_str() {
            "equalitycomparer" => "EqualityComparer",
            "stringcomparer" => "StringComparer",
            _ => "Comparer",
        });
    }
    if class.eq_ignore_ascii_case("Nullable")
        && method_name.eq_ignore_ascii_case("GetUnderlyingType")
    {
        return Some("Type");
    }
    if class.eq_ignore_ascii_case("Object")
        && matches!(
            method_name.to_ascii_lowercase().as_str(),
            "equals" | "referenceequals"
        )
    {
        return Some("Boolean");
    }
    if class.eq_ignore_ascii_case("TimeSpan")
        && matches!(
            method_name.to_ascii_lowercase().as_str(),
            "fromdays"
                | "fromhours"
                | "fromminutes"
                | "fromseconds"
                | "frommilliseconds"
                | "parse"
                | "zero"
                | "add"
                | "subtract"
        )
    {
        return Some("TimeSpan");
    }
    if class.eq_ignore_ascii_case("TimeSpan") && method_name.eq_ignore_ascii_case("Compare") {
        return Some("Int32");
    }
    if class.eq_ignore_ascii_case("DateTime")
        && matches!(
            method_name.to_ascii_lowercase().as_str(),
            "now"
                | "utcnow"
                | "today"
                | "minvalue"
                | "maxvalue"
                | "parse"
                | "parseexact"
                | "frombinary"
                | "fromfiletimeutc"
                | "fromoadate"
        )
    {
        return Some("DateTime");
    }
    if class.eq_ignore_ascii_case("DateTime") {
        return match method_name.to_ascii_lowercase().as_str() {
            "compare" => Some("Int32"),
            "equals" | "isleapyear" | "tryparse" | "tryparseexact" => Some("Boolean"),
            _ => None,
        };
    }
    // `System.TimeZoneInfo`. Without these the local in
    // `Dim ny = TimeZoneInfo.FindSystemTimeZoneById(…)` carried no type, so
    // `ny.GetUtcOffset(dt)` had nothing to resolve the INSTANCE member against
    // and answered `undefined` — the members were registered, what they
    // RETURN was not.
    if class.eq_ignore_ascii_case("TimeZoneInfo") {
        return match method_name.to_ascii_lowercase().as_str() {
            "utc" | "local" | "findsystemtimezonebyid" => Some("TimeZoneInfo"),
            "converttime" | "converttimetoutc" | "converttimefromutc" => Some("DateTime"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("DateTimeOffset")
        && matches!(
            method_name.to_ascii_lowercase().as_str(),
            "now" | "utcnow" | "parse" | "fromunixtimeseconds" | "fromunixtimemilliseconds"
        )
    {
        return Some("DateTimeOffset");
    }
    if class.eq_ignore_ascii_case("DateTimeOffset") {
        return match method_name.to_ascii_lowercase().as_str() {
            "compare" => Some("Int32"),
            "equals" | "tryparse" => Some("Boolean"),
            "minvalue" | "maxvalue" => Some("DateTimeOffset"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Stopwatch") {
        return match method_name.to_ascii_lowercase().as_str() {
            "startnew" => Some("Stopwatch"),
            "gettimestamp" | "frequency" => Some("Int64"),
            "ishighresolution" => Some("Boolean"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("BitConverter") {
        return match method_name.to_ascii_lowercase().as_str() {
            "tochar" => Some("Char"),
            "toboolean" => Some("Boolean"),
            "todouble" | "tosingle" => Some("Double"),
            "toint16" | "toint32" | "touint16" | "touint32" => Some("Int32"),
            "toint64" | "touint64" => Some("Int64"),
            "tostring" => Some("String"),
            "getbytes" => Some("Array"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Convert") && method_name.eq_ignore_ascii_case("ToDateTime") {
        return Some("DateTime");
    }
    if class.eq_ignore_ascii_case("File") {
        return match method_name.to_ascii_lowercase().as_str() {
            "readalltext" => Some("String"),
            "readalllines" | "readallbytes" => Some("Array"),
            "create" | "openread" => Some("FileStream"),
            "exists" => Some("Boolean"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Directory") {
        return match method_name.to_ascii_lowercase().as_str() {
            "getfiles" | "getdirectories" => Some("Array"),
            "exists" => Some("Boolean"),
            "getcurrentdirectory" => Some("String"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Path") {
        return match method_name.to_ascii_lowercase().as_str() {
            "combine"
            | "gettemppath"
            | "gettempfilename"
            | "getfilename"
            | "getfilenamewithoutextension"
            | "getextension"
            | "getdirectoryname"
            | "getfullpath"
            | "getpathroot"
            | "changeextension"
            | "trimendingdirectoryseparator" => Some("String"),
            "ispathrooted" | "hasextension" | "endsindirectoryseparator" => Some("Boolean"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Uri") {
        return match method_name.to_ascii_lowercase().as_str() {
            "trycreate" | "makerelativeuri" => Some("Uri"),
            "isbaseof" | "iswellformeduristring" => Some("Boolean"),
            "escapedatastring" | "unescapedatastring" => Some("String"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Guid")
        && matches!(
            method_name.to_ascii_lowercase().as_str(),
            "empty" | "newguid" | "parse"
        )
    {
        return Some("Guid");
    }
    if class.eq_ignore_ascii_case("Version") && method_name.eq_ignore_ascii_case("Parse") {
        return Some("Version");
    }
    if class.eq_ignore_ascii_case("Version") {
        return match method_name.to_ascii_lowercase().as_str() {
            "compareto" => Some("Int32"),
            "equals" => Some("Boolean"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Process") {
        return match method_name.to_ascii_lowercase().as_str() {
            "getcurrentprocess" | "getprocessbyid" | "start" => Some("Process"),
            "getprocesses" | "getprocessesbyname" => Some("Array"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("CancellationToken") && method_name.eq_ignore_ascii_case("None") {
        return Some("CancellationToken");
    }
    if class.eq_ignore_ascii_case("CancellationTokenSource")
        && method_name.eq_ignore_ascii_case("CreateLinkedTokenSource")
    {
        return Some("CancellationTokenSource");
    }
    if class.eq_ignore_ascii_case("Task")
        && matches!(
            method_name.to_ascii_lowercase().as_str(),
            "delay" | "yield" | "whenall" | "whenany" | "fromresult" | "run"
        )
    {
        return Some("Task");
    }
    if class.eq_ignore_ascii_case("Encoding")
        && matches!(
            method_name.to_ascii_lowercase().as_str(),
            "utf8"
                | "ascii"
                | "unicode"
                | "default"
                | "utf32"
                | "latin1"
                | "bigendianunicode"
                | "getencoding"
        )
    {
        return Some("Encoding");
    }
    None
}

pub fn static_property_type(class_name: &str, property_name: &str) -> Option<&'static str> {
    let class = class_name.rsplit('.').next().unwrap_or(class_name);
    if class.eq_ignore_ascii_case("DateTime")
        && matches!(
            property_name.to_ascii_lowercase().as_str(),
            "now" | "utcnow" | "today" | "minvalue" | "maxvalue"
        )
    {
        return Some("DateTime");
    }
    if class.eq_ignore_ascii_case("DateTimeOffset")
        && matches!(
            property_name.to_ascii_lowercase().as_str(),
            "now" | "utcnow" | "minvalue" | "maxvalue"
        )
    {
        return Some("DateTimeOffset");
    }
    if class.eq_ignore_ascii_case("DateTimeKind")
        && matches!(
            property_name.to_ascii_lowercase().as_str(),
            "utc" | "local" | "unspecified"
        )
    {
        return Some("String");
    }
    if class.eq_ignore_ascii_case("TimeSpan") && property_name.eq_ignore_ascii_case("Zero") {
        return Some("TimeSpan");
    }
    if class.eq_ignore_ascii_case("Guid") && property_name.eq_ignore_ascii_case("Empty") {
        return Some("Guid");
    }
    if class.eq_ignore_ascii_case("Stopwatch") {
        return match property_name.to_ascii_lowercase().as_str() {
            "frequency" => Some("Int64"),
            "ishighresolution" => Some("Boolean"),
            _ => None,
        };
    }
    None
}

/// Properties whose TYPE the platform declares, enumerable so tree
/// registration can publish them.
///
/// A property read had no declared type at all — the registrar's scan walks
/// `class.methods` only — so the value came back untyped and the next hop on it
/// resolved against nothing. `cmd.Parameters` is the case that matters:
/// `wasi:sql` has always had `params.add-with-value`/`clear`/`count`, and
/// `make_command_obj` puts a params object on every command, but with no type
/// on `Parameters` the chain died at `.AddWithValue(…)` and the query ran with
/// no parameters bound — "Got 0, needed 2".
///
/// Enumerable rather than a lookup so there is ONE declaration serving both the
/// walker's type inference and the tree.
pub fn declared_instance_property_types(
    class_name: &str,
) -> &'static [(&'static str, &'static str)] {
    match class_name
        .rsplit('.')
        .next()
        .unwrap_or(class_name)
        .to_ascii_lowercase()
        .as_str()
    {
        "sqlcommand" | "oledbcommand" | "adodbcommand" => {
            &[("Parameters", "SqlParameterCollection")]
        }
        // The cursor's own state. These are plain struct fields the ADODB
        // adapter and `wasi:sql` already write — what was missing is any
        // statement of WHAT they are, and a frontend that has to decide an
        // operator's meaning from the operand's type cannot decide without one.
        //
        // VB's `Not` is bitwise on a number and logical on a Boolean, chosen at
        // COMPILE time from the inferred type. With `EOF` undeclared,
        // `Do While Not rs.EOF` took the bitwise arm — `Not False` is `-1` and
        // `Not True` is `-2`, both truthy — so the loop walked the real rows and
        // then spun forever on empty ones.
        // Only the BOOLEANS. `RecordCount`/`FieldCount`/`Position`/`Count` are
        // every one of them stored as an `f64` — `collections::emit_len` yields
        // `ecma:array.length`, `wasi:sql` writes `Value::F64(cols.len())`, and
        // this adapter's own constructors write `Value::F64(0.0)`. Declaring
        // them `Int32` would state a width the storage does not have, which is
        // the same type-vs-storage split that made `EOF` loop forever, pointed
        // the other way. Nothing needs them declared: VB's `Not` is bitwise on
        // a numeric whether or not the type is known, so the declaration only
        // ever mattered for the Booleans.
        "recordset" => &[("EOF", "Boolean"), ("IsClosed", "Boolean")],
        "sqldatareader" | "oledbdatareader" | "datareader" => {
            &[("HasRows", "Boolean"), ("IsClosed", "Boolean")]
        }
        "timezoneinfo" => &[
            ("Id", "String"),
            ("StandardName", "String"),
            ("DisplayName", "String"),
            ("BaseUtcOffset", "TimeSpan"),
        ],
        // `DateTimeOffset`'s composite properties. Undeclared, `dto.Date` read
        // back untyped and the NEXT hop resolved against nothing —
        // `dto.Date.ToString("yyyy-MM-dd")` rendered `[object datetime]`
        // because the 1-arg `ToString` could not be found on an unknown type.
        "datetimeoffset" => &[
            ("Date", "DateTime"),
            ("DateTime", "DateTime"),
            ("UtcDateTime", "DateTime"),
            ("Offset", "TimeSpan"),
            ("TimeOfDay", "TimeSpan"),
        ],
        // The same for `DateTime` itself.
        "datetime" => &[("Date", "DateTime"), ("TimeOfDay", "TimeSpan")],
        // `tcs.Task` hands back the consumer half. Undeclared, the promise read
        // back untyped and the NEXT hop resolved against nothing —
        // `tcs.Task.IsCompleted` and `tcs.Task.Status` answered `undefined`
        // even though `Task` declares both, because the receiver never said
        // what it was.
        "taskcompletionsource" => &[("Task", "Task")],
        _ => &[],
    }
}

pub fn instance_property_type(class_name: &str, property_name: &str) -> Option<&'static str> {
    let class = class_name.rsplit('.').next().unwrap_or(class_name);
    // `cmd.Parameters` — the collection the host already puts on every command
    // object (`make_command_obj` inserts `parameters`). Saying WHAT it is is
    // what lets `cmd.Parameters.AddWithValue(...)` resolve; without the type
    // the chain died at the second hop, the parameters never reached the
    // query, and `wasi:sql` answered "Got 0, needed 2".
    if let Some((_, ty)) = declared_instance_property_types(class)
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(property_name))
    {
        return Some(ty);
    }
    if class.eq_ignore_ascii_case("StringBuilder") {
        return match property_name.to_ascii_lowercase().as_str() {
            "length" | "capacity" | "maxcapacity" => Some("Int32"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Process") {
        return match property_name.to_ascii_lowercase().as_str() {
            "hasexited" => Some("Boolean"),
            "id" | "handle" => Some("Int32"),
            "processname" | "priorityclass" => Some("String"),
            "workingset64" | "peakworkingset64" | "virtualmemorysize64" | "privatememorysize64" => {
                Some("Int64")
            }
            "starttime" => Some("DateTime"),
            "totalprocessortime" | "userprocessortime" => Some("TimeSpan"),
            "threads" | "modules" | "mainmodule" => Some("Object"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("TimeSpan") {
        return match property_name.to_ascii_lowercase().as_str() {
            "ticks" => Some("Int64"),
            "totalmilliseconds" | "totalseconds" | "totalminutes" | "totalhours" | "totaldays" => {
                Some("Double")
            }
            "days" | "hours" | "minutes" | "seconds" => Some("Int32"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("DateTime") {
        return match property_name.to_ascii_lowercase().as_str() {
            "year" | "month" | "day" | "hour" | "minute" | "second" | "millisecond"
            | "dayofyear" => Some("Int32"),
            "ticks" => Some("Int64"),
            "date" => Some("DateTime"),
            "timeofday" => Some("TimeSpan"),
            "kind" | "dayofweek" => Some("String"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("DateTimeOffset") {
        return match property_name.to_ascii_lowercase().as_str() {
            "year" | "month" | "day" | "hour" | "minute" | "second" | "millisecond"
            | "dayofyear" => Some("Int32"),
            "ticks" => Some("Int64"),
            "offset" => Some("TimeSpan"),
            "datetime" | "utcdatetime" | "date" => Some("DateTime"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Stopwatch") {
        return match property_name.to_ascii_lowercase().as_str() {
            "isrunning" => Some("Boolean"),
            "elapsedmilliseconds" | "elapsedticks" => Some("Int64"),
            "elapsed" => Some("TimeSpan"),
            _ => None,
        };
    }
    if class.eq_ignore_ascii_case("Task") {
        return match property_name.to_ascii_lowercase().as_str() {
            "iscompleted" | "iscanceled" => Some("Boolean"),
            "result" => Some("Object"),
            _ => None,
        };
    }
    None
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
            ConstructorTarget::Common("dotnet.list_new".to_string())
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
                emit: "dotnet.console_writeline".to_string()
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
        assert!(is_noop_method("SuspendLayout"));
        assert!(is_known_constant("PI"));
    }
}

// ── Registry shims ──────────────────────────────────────────────────────
//
// Free functions over `surface()` so the compiler can reach these through
// `PlatformDef` function pointers instead of naming this crate.

pub fn registry_lookup_constructor(
    name: &str,
) -> Option<vybe_runtime::component_model::ConstructorTarget> {
    surface().lookup_constructor(name)
}

pub fn registry_lookup_instance_method(
    class_name: &str,
    method_name: &str,
    arg_count: u8,
) -> Option<vybe_runtime::component_model::InstanceMethodTarget> {
    surface().lookup_instance_method(class_name, method_name, arg_count)
}

pub fn registry_lookup_instance_property(
    class_name: &str,
    property_name: &str,
) -> Option<vybe_runtime::component_model::InstancePropertyTarget> {
    surface().lookup_instance_property(class_name, property_name)
}

pub fn registry_is_known_constant(name: &str) -> bool {
    surface().is_known_constant(name)
}
