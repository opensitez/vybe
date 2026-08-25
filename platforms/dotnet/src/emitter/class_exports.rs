use std::sync::LazyLock;

use vybe_runtime::component_model::ClassType;

use super::winforms::classes::DotnetClass;

#[derive(Debug, Clone)]
pub struct DotnetClassExport {
    pub interface: &'static str,
    pub class: ClassType,
    pub wrapper: Option<DotnetClass>,
}

impl DotnetClassExport {
    pub fn new(interface: &'static str, class: ClassType) -> Self {
        Self {
            interface,
            class,
            wrapper: None,
        }
    }

    pub fn with_wrapper(interface: &'static str, class: ClassType, wrapper: DotnetClass) -> Self {
        Self {
            interface,
            class,
            wrapper: Some(wrapper),
        }
    }
}

pub fn dotnet_class_exports() -> &'static [DotnetClassExport] {
    static EXPORTS: LazyLock<Vec<DotnetClassExport>> = LazyLock::new(|| {
        let mut exports = Vec::new();
        exports.extend_from_slice(super::core::component_classes::class_exports());
        exports.extend_from_slice(super::winforms::component_classes::class_exports());
        add_keyword_aliases(&mut exports);
        exports
    });
    EXPORTS.as_slice()
}

/// The C# predefined-type keywords, declared as the types they name.
///
/// ⛔ `object`, `string`, `int` … are NOT case variants of `Object`, `String`,
/// `Int32`. They are distinct C# spellings for the same type — the same
/// category as the `Byte`/`byte` pair that `component_classes_system_values`
/// already declares in both spellings. While the namespace tree folded every
/// lookup these resolved by accident; once the fold became conditional on the
/// language's directive, C# — which does not fold — stopped finding them and
/// `new object()` / `string.Format(…)` reached `undefined is not callable`.
///
/// Declared by CLONING the CLR type's export under the alias spelling, so the
/// two can never drift: there is one definition and two names for it.
///
/// VB's `Integer`/`Long`/`Single` spellings are deliberately absent — VB
/// reaches `canonical_type_name` before the tree, and adding them here would
/// give one fact two homes.
fn add_keyword_aliases(exports: &mut Vec<DotnetClassExport>) {
    const ALIASES: &[(&str, &str)] = &[
        ("Object", "object"),
        ("String", "string"),
        ("Boolean", "bool"),
        ("Int16", "short"),
        ("UInt16", "ushort"),
        ("Int32", "int"),
        ("UInt32", "uint"),
        ("Int64", "long"),
        ("UInt64", "ulong"),
        ("Single", "float"),
    ];
    let mut aliased = Vec::new();
    for (clr, keyword) in ALIASES {
        // The FIRST export declaring the CLR name wins, matching the lookup.
        if let Some(source) = exports
            .iter()
            .find(|e| e.class.name.eq_ignore_ascii_case(clr))
        {
            // Already declared in both spellings (the numeric parse classes
            // do this themselves) — nothing to add.
            if exports.iter().any(|e| e.class.name == *keyword) {
                continue;
            }
            let mut clone = source.clone();
            clone.class.name = (*keyword).to_string();
            aliased.push(clone);
        }
    }
    exports.extend(aliased);
}
