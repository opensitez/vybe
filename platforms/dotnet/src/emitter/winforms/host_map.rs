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
        "system.windows.forms" | "application" => Some("vybe:gui"),
        _ => None,
    }
}

pub fn map_host_func(module: &str, func: &str) -> Option<String> {
    match (module, func) {
        ("vybe:gui", f) => {
            let canonical = vybe_compiler::primitives::gui::canonical_control_name(f);
            if !canonical.is_empty() && canonical != f {
                Some(vybe_compiler::primitives::gui::host_fn_new_control(
                    &canonical,
                ))
            } else {
                Some(f.to_string())
            }
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_static_method_mappings_exclude_common_application_adapters() {
        assert!(static_method_mappings()
            .iter()
            .all(|mapping| mapping.type_name != "Application"));
    }
}
