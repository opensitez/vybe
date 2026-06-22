use std::sync::LazyLock;

use super::super::types::{KnownTypeMapping, KnownTypeTarget};
use vybe_bytecode::component_model::ConstructorTarget;

const KNOWN_CONSTANTS: &[&str] = &[
    "pi",
    "e",
    "maxvalue",
    "minvalue",
    "positiveinfinity",
    "negativeinfinity",
    "nan",
    "epsilon",
    "empty",
    "newline",
    "true",
    "false",
    "completedtask",
];

static KNOWN_TYPE_MAPPINGS: LazyLock<Vec<KnownTypeMapping>> = LazyLock::new(|| {
    super::component_classes::class_exports()
        .iter()
        .filter_map(|export| {
            let target = export.class.constructor.as_ref()?.backing.as_ref()?;
            Some(KnownTypeMapping {
                name: leak_string(export.class.name.to_lowercase()),
                interface: export.interface,
                display_name: leak_string(export.class.name.clone()),
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

pub fn is_known_constant(name: &str) -> bool {
    known_constants().contains(&name)
}

pub fn known_constants() -> &'static [&'static str] {
    KNOWN_CONSTANTS
}

pub fn capitalize_data_type(name: &str) -> String {
    match name {
        "dataset" => "DataSet",
        "datatable" => "DataTable",
        "dataadapter" => "DataAdapter",
        _ => return String::new(),
    }
    .to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_known_type_mappings_exclude_winforms_entries() {
        assert!(
            known_type_mappings()
                .iter()
                .any(|mapping| mapping.name == "stringbuilder")
        );
        assert!(
            !known_type_mappings()
                .iter()
                .any(|mapping| mapping.name == "form")
        );
    }
}
