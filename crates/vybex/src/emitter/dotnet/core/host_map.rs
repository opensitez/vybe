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
        "system.math" => Some("vybe:math"),
        "system.convert" => Some("vybe:convert"),
        "system.string" => Some("vybe:string"),
        "system.array" => Some("vybe:array"),
        "system.environment" => Some("wasi:cli"),
        "system.io.streamreader" | "system.io.streamwriter" => Some("dotnet:io"),
        "system.io" | "system.io.file" | "system.io.path" | "system.io.directory" => Some("wasi:filesystem"),
        "system.threading.thread" => Some("wasi:clocks"),
        "system.threading" | "system.threading.tasks" => Some("vybe:threading"),
        "system.diagnostics.process" => Some("vybe:types"),
        "system.diagnostics.stopwatch" => Some("wasi:clocks"),
        "system.diagnostics.debug" | "system.diagnostics.trace" | "system.diagnostics" => Some("wasi:cli"),
        "system.net" => Some("wasi:http"),
        "system.net.dns" => Some("dotnet:net"),
        "system.net.sockets" => Some("dotnet:sockets"),
        "system.text.regularexpressions" => Some("vybe:regex"),
        "system.text" => Some("vybe:string"),
        "system.collections.generic" | "system.collections" => Some("vybe:types"),
        "system.data" | "system.data.sqlclient" | "system.data.oledb" => Some("vybe:data"),
        "system.security.cryptography" => Some("vybe:crypto"),
        "system.xml.linq" => Some("vybe:xml"),
        "system.drawing" => Some("vybe:drawing"),
        "microsoft.visualbasic" => Some("vybe:string"),
        _ => None,
    }
}

pub fn map_host_func(module: &str, func: &str) -> Option<String> {
    match (module, func) {
        ("vybe:math", f) => Some(f.to_string()),
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
}