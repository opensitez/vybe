use super::super::types::{KnownTypeMapping, KnownTypeTarget};
use vybe_bytecode::component_model::ConstructorTarget;

const NOOP_METHODS: &[&str] = &[
    "suspendlayout", "resumelayout", "performlayout",
    "refresh", "invalidate", "update", "begininit", "endinit",
    "dispose", "select", "focus", "bringtofront", "sendtoback",
    "createcontrol", "show", "hide",
];

const NAMESPACE_CONSTANTS: &[(&str, f64)] = &[
    ("autoscalemode.none", 0.0),
    ("autoscalemode.font", 1.0),
    ("autoscalemode.dpi", 2.0),
    ("autoscalemode.inherit", 3.0),
    ("autosizemode.growonly", 0.0),
    ("autosizemode.growandshrink", 1.0),
    ("formstartposition.manual", 0.0),
    ("formstartposition.centerscreen", 1.0),
    ("formstartposition.windowsdefaultlocation", 2.0),
    ("formstartposition.windowsdefaultbounds", 3.0),
    ("formstartposition.centerparent", 4.0),
    ("formwindowstate.normal", 0.0),
    ("formwindowstate.minimized", 1.0),
    ("formwindowstate.maximized", 2.0),
    ("dialogresult.none", 0.0),
    ("dialogresult.ok", 1.0),
    ("dialogresult.cancel", 2.0),
    ("dialogresult.abort", 3.0),
    ("dialogresult.retry", 4.0),
    ("dialogresult.ignore", 5.0),
    ("dialogresult.yes", 6.0),
    ("dialogresult.no", 7.0),
    ("messageboxbuttons.ok", 0.0),
    ("messageboxbuttons.okcancel", 1.0),
    ("messageboxbuttons.abortretryignore", 2.0),
    ("messageboxbuttons.yesnocancel", 3.0),
    ("messageboxbuttons.yesno", 4.0),
    ("messageboxbuttons.retrycancel", 5.0),
    ("messageboxbuttons.canceltrycontinue", 6.0),
    ("messageboxicon.none", 0.0),
    ("messageboxicon.error", 16.0),
    ("messageboxicon.question", 32.0),
    ("messageboxicon.warning", 48.0),
    ("messageboxicon.information", 64.0),
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

pub fn namespace_constants() -> &'static [(&'static str, f64)] {
    NAMESPACE_CONSTANTS
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