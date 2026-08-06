use std::sync::LazyLock;

use super::super::host_map::DotnetStaticMethodMapping;
use vybe_runtime::component_model::MethodBody;

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
        "system.console" => Some("web:console"),
        "system.math" => Some("ecma:math"),
        // System.Convert.* lowers via the emitter's convert opcodes;
        // fall through to the type registry / walker rewrites.
        "system.convert" => None,
        // `system.string.<X>` falls back here when not matched by an
        // explicit entry in `component_classes.rs::STATIC_ONLY_CLASS`.
        // Routing through `ecma:string` means .NET callers get the ECMA-262
        // String surface; .NET-only methods (Format, IsNullOrEmpty, Join with
        // (separator, array) signature) need explicit adapter entries.
        "system.string" => Some("ecma:string"),
        // `system.array.X` falls back here when not matched by an
        // explicit entry in `component_classes.rs::STATIC_ONLY_CLASS`.
        // Routing through `ecma:array` means .NET callers get the
        // ECMA-262 Array surface by default; .NET-only methods need
        // explicit adapter entries.
        "system.array" => Some("ecma:array"),
        // System.Environment has explicit component metadata and common
        // adapters; do not synthesize `wasi:cli/environment.*` fallbacks.
        "system.environment" => None,
        // StreamReader / StreamWriter — retired from `dotnet:io` host
        // namespace; lower at compile time through
        // `emitter::dotnet::core::stream_io_adapter` (composes
        // `node:fs.{read,write}FileSync`). The mapping returns None so
        // FQN resolution falls through to the explicit class entries in
        // `component_classes.rs`.
        "system.io" | "system.io.file" | "system.io.path" | "system.io.directory" => {
            Some("wasi:filesystem")
        }
        "system.threading.thread" => None,
        // System.Threading / Tasks: spawn+join compile to Op::THREAD_SPAWN /
        // Op::THREAD_JOIN; Thread.Sleep compiles via thread_adapter to
        // wasi:clocks/monotonic-clock + wasi:io/poll. Return None so the FQN
        // resolver falls through to the type registry.
        "system.diagnostics.process" => None,
        "system.diagnostics.stopwatch" => Some("wasi:clocks/monotonic-clock"),
        "system.diagnostics.debug" | "system.diagnostics.trace" | "system.diagnostics" => {
            Some("web:console")
        }
        // System.Net has no single backing host module: `wasi:http` is a WASI
        // *package*, not an interface (the interfaces are `wasi:http/types`,
        // `wasi:http/outgoing-handler`, `wasi:http/incoming-handler`), and the
        // spec surface is resource-based with no one-call `fetch`. So the
        // classes lower through `http_adapter`, which emits the real
        // request -> outgoing-handler.handle -> consume-body sequence. Return
        // None so the FQN resolver falls through to the type registry.
        "system.net" => None,
        // Networking no longer falls through to a monolithic `.NET`
        // host namespace. `System.Net.Dns` and
        // `System.Net.Sockets.*` resolve through explicit component
        // metadata and lower via emitter adapters that compose
        // `wasi:sockets/*` + `node:os`.
        // .NET `System.Text.RegularExpressions.Regex.*` falls back here.
        // Same pattern-first arg shape as PHP/Python, but ecma:regexp.test
        // and exec already accept string-as-pattern. Methods needing arg
        // reorder (Replace/Split with input-first .NET shape) get explicit
        // adapter entries via `STATIC_ONLY_CLASS` (or stdlib chunks).
        "system.text.regularexpressions" => Some("ecma:regexp"),
        "system.text" => Some("ecma:string"),
        "system.collections.generic" | "system.collections" => None,
        "system.data.sqlclient" => Some("wasi:sql/types"),
        "system.data.oledb" => Some("wasi:sql/types"),
        "adodb" => Some("wasi:sql/types"),
        // DataTable/DataSet constructors lower through datatable_adapter.rs;
        // method dispatch is handled by DotnetClassExport bindings. Fall through.
        "system.data" => None,
        "system.security.cryptography" => Some("wasi:crypto/hashes"),
        "system.xml.linq" => Some("web:dom-parser"),
        "system.drawing" => Some("vybe:gui"),
        "microsoft.visualbasic" => None,
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
        assert!(static_method_mappings()
            .iter()
            .any(|mapping| mapping.type_name == "Convert"));
        assert!(!static_method_mappings()
            .iter()
            .any(|mapping| mapping.type_name == "Application"));
    }

    #[test]
    fn test_network_namespaces_do_not_fall_back_to_retired_dotnet_host_modules() {
        assert_eq!(namespace_to_host_module("system.net"), None);
        assert_eq!(namespace_to_host_module("system.net.dns"), None);
        assert_eq!(namespace_to_host_module("system.net.sockets"), None);
    }
}
