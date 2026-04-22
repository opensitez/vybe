use super::super::types::{KnownTypeMapping, KnownTypeTarget};
use vybe_bytecode::component_model::ConstructorTarget;

const NOOP_METHODS: &[&str] = &[
    "suspendlayout", "resumelayout", "performlayout",
    "refresh", "invalidate", "update", "begininit", "endinit",
    "dispose", "select", "focus", "bringtofront", "sendtoback",
    "createcontrol", "show", "hide",
];

use std::sync::LazyLock;

static KNOWN_TYPE_MAPPINGS: LazyLock<Vec<KnownTypeMapping>> = LazyLock::new(|| {
    super::component_classes::class_exports()
        .iter()
        .filter_map(|export| {
            if export.class.name != "Form" {
                return None;
            }
            let target = export.class.constructor.as_ref()?.backing.as_ref()?;
            Some(KnownTypeMapping {
                name: "form",
                interface: export.interface,
                display_name: "Form",
                target: match target {
                    ConstructorTarget::Host(target) => KnownTypeTarget::Host {
                        module: leak_string(target.module.clone()),
                        constructor: leak_string(target.name.clone()),
                    },
                    ConstructorTarget::Common(name) => KnownTypeTarget::Common {
                        emit: leak_string(name.clone()),
                    },
                },
            })
        })
        .collect()
});

fn leak_string(value: String) -> &'static str {
    Box::leak(value.into_boxed_str())
}

pub fn known_type_mappings() -> &'static [KnownTypeMapping] {
    KNOWN_TYPE_MAPPINGS.as_slice()
}

pub fn is_noop_method(name: &str) -> bool {
    noop_methods().contains(&name)
}

pub fn noop_methods() -> &'static [&'static str] {
    NOOP_METHODS
}

pub fn capitalize_control_name(name: &str) -> String {
    crate::emitter::gui::canonical_control_name(name)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_known_type_mappings_include_form() {
        assert!(known_type_mappings().iter().any(|mapping| mapping.name == "form"));
    }
}