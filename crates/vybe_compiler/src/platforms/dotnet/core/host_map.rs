use std::sync::LazyLock;

use super::super::host_map::DotnetStaticMethodMapping;
use vybe_bytecode::component_model::MethodBody;

static STATIC_METHOD_MAPPINGS: LazyLock<Vec<DotnetStaticMethodMapping>> = LazyLock::new(|| {
    super::component_classes::class_exports()
        .iter()
        .flat_map(|export| {
            export.class.methods.iter().filter_map(move |method| {
                if !method.is_static {
                    return None;
                }
                let MethodBody::HostCall(target) = &method.body else {
                    return None;
                };
                Some(DotnetStaticMethodMapping {
                    interface: export.interface,
                    type_name: leak_string(export.class.name.clone()),
                    method_name: leak_string(method.name.clone()),
                    host_module: leak_string(target.module.clone()),
                    host_fn: leak_string(target.name.clone()),
                    arity: method.arity,
                })
            })
        })
        .collect()
});

fn leak_string(value: String) -> &'static str {
    Box::leak(value.into_boxed_str())
}

pub fn static_method_mappings() -> &'static [DotnetStaticMethodMapping] {
    STATIC_METHOD_MAPPINGS.as_slice()
}

pub fn namespace_to_host_module(prefix: &str) -> Option<&'static str> {
    match prefix {
        "system.console" => Some("wasi:cli"),
        "system.math" => Some("ecma:math"),
        "system.convert" => Some("vybe:convert"),
        // `system.string.<X>` falls back here when not matched by an
        // explicit entry in `component_classes.rs::STATIC_ONLY_CLASS`
        // (e.g. `String.Format` → `vybe:string.format`). Routing through
        // `ecma:string` means .NET callers get the ECMA-262 String surface
        // by default; .NET-only methods (Format, IsNullOrEmpty, Join with
        // (separator, array) signature) need explicit adapter entries.
        "system.string" => Some("ecma:string"),
        // `system.array.X` falls back here when not matched by an
        // explicit entry in `component_classes.rs::STATIC_ONLY_CLASS`.
        // Routing through `ecma:array` means .NET callers get the
        // ECMA-262 Array surface by default; .NET-only methods need
        // explicit adapter entries.
        "system.array" => Some("ecma:array"),
        "system.environment" => Some("wasi:cli"),
        // StreamReader / StreamWriter — retired from `dotnet:io` host
        // namespace; lower at compile time through
        // `emitter::dotnet::core::stream_io_adapter` (composes
        // `node:fs.{read,write}FileSync`). The mapping returns None so
        // FQN resolution falls through to the explicit class entries in
        // `component_classes.rs`.
        "system.io" | "system.io.file" | "system.io.path" | "system.io.directory" => Some("wasi:filesystem"),
        "system.threading.thread" => Some("wasi:clocks"),
        // System.Threading / Tasks: spawn+join compile to Op::THREAD_SPAWN /
        // Op::THREAD_JOIN; sleep uses wasi:clocks; Task/Thread class types
        // are registered in the TypeRegistry but carry no host-fn namespace.
        // Return None so the FQN resolver falls through to the type registry
        // rather than pointing at a dead `vybe:threading` module.
        "system.diagnostics.process" => Some("vybe:types"),
        "system.diagnostics.stopwatch" => Some("wasi:clocks"),
        "system.diagnostics.debug" | "system.diagnostics.trace" | "system.diagnostics" => Some("wasi:cli"),
        "system.net" => Some("wasi:http"),
        // Networking no longer falls through to a monolithic `.NET`
        // host namespace. `System.Net.Http` maps directly to
        // `wasi:http`, while `System.Net.Dns` and
        // `System.Net.Sockets.*` resolve through explicit component
        // metadata and lower via emitter adapters that compose
        // `wasi:sockets/*` + `node:os`.
        // .NET `System.Text.RegularExpressions.Regex.*` falls back here.
        // Same pattern-first arg shape as PHP/Python, but ecma:regexp.test
        // and exec already accept string-as-pattern. Methods needing arg
        // reorder (Replace/Split with input-first .NET shape) get explicit
        // adapter entries via `STATIC_ONLY_CLASS` (or stdlib chunks).
        "system.text.regularexpressions" => Some("ecma:regexp"),
        "system.text" => Some("vybe:string"),
        "system.collections.generic" | "system.collections" => Some("vybe:types"),
        "system.data.sqlclient" => Some("wasi:sql/types"),
        "system.data.oledb" => Some("wasi:sql/types"),
        "adodb" => Some("wasi:sql/types"),
        "system.data" => Some("vybe:data"),
        "system.security.cryptography" => Some("vybe:crypto"),
        "system.xml.linq" => Some("vybe:xml"),
        "system.drawing" => Some("vybe:drawing"),
        "microsoft.visualbasic" => Some("vybe:string"),
        _ => None,
    }
}

pub fn map_host_func(module: &str, func: &str) -> Option<String> {
    match (module, func) {
        ("ecma:math", f) => Some(f.to_string()),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_static_method_mappings_exclude_winforms_application() {
        assert!(static_method_mappings().iter().any(|mapping| mapping.type_name == "Console"));
        assert!(!static_method_mappings().iter().any(|mapping| mapping.type_name == "Application"));
    }

    #[test]
    fn test_network_namespaces_do_not_fall_back_to_retired_dotnet_host_modules() {
        assert_eq!(namespace_to_host_module("system.net"), Some("wasi:http"));
        assert_eq!(namespace_to_host_module("system.net.dns"), None);
        assert_eq!(namespace_to_host_module("system.net.sockets"), None);
    }
}