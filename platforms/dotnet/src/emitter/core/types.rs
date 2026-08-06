use std::sync::LazyLock;

use super::super::types::{KnownTypeMapping, KnownTypeTarget};
use vybe_runtime::component_model::ConstructorTarget;

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

// π, e and τ are NOT part of this platform's surface — they are numbers, and
// `vybe_compiler::primitives::math` owns them for every language. The twelve
// rows that used to sit at the top of this table (`math.pi`, `Math.PI`,
// `system.math.pi`, `System.Math.PI`, ×3 concepts) made the dotnet platform the
// de-facto owner of `Math.PI`, so any language that wanted π had to reach a
// platform it has nothing to do with. `math::dotted_constant` matches the owner
// segment case-insensitively and so answers all four spellings on its own.
//
// What REMAINS here is genuinely .NET's: type limits spelled by .NET type names
// (`int.MaxValue`), and the framework enum ordinals (`CommandType`,
// `ConnectionState`, `RegexOptions`, `MsgBoxStyle`).
const NAMESPACE_CONSTANTS: &[(&str, f64)] = &[
    ("int.MaxValue", 2_147_483_647.0),
    ("int.MinValue", -2_147_483_648.0),
    ("double.MaxValue", f64::MAX),
    ("double.MinValue", -f64::MAX),
    ("double.NaN", f64::NAN),
    ("double.PositiveInfinity", f64::INFINITY),
    ("double.NegativeInfinity", f64::NEG_INFINITY),
    ("float.MaxValue", 3.4028235e38),
    ("float.MinValue", -3.4028235e38),
    ("char.MaxValue", 65535.0),
    ("char.MinValue", 0.0),
    ("commandtype.text", 1.0),
    ("CommandType.Text", 1.0),
    ("commandtype.storedprocedure", 4.0),
    ("CommandType.StoredProcedure", 4.0),
    ("connectionstate.closed", 0.0),
    ("ConnectionState.Closed", 0.0),
    ("connectionstate.open", 1.0),
    ("ConnectionState.Open", 1.0),
    ("regexoptions.none", 0.0),
    ("RegexOptions.None", 0.0),
    ("system.text.regularexpressions.regexoptions.none", 0.0),
    ("System.Text.RegularExpressions.RegexOptions.None", 0.0),
    ("regexoptions.ignorecase", 1.0),
    ("RegexOptions.IgnoreCase", 1.0),
    ("msgboxstyle.okonly", 0.0),
    ("MsgBoxStyle.OkOnly", 0.0),
    ("microsoft.visualbasic.msgboxstyle.okonly", 0.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.OkOnly", 0.0),
    ("msgboxstyle.okcancel", 1.0),
    ("MsgBoxStyle.OkCancel", 1.0),
    ("microsoft.visualbasic.msgboxstyle.okcancel", 1.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.OkCancel", 1.0),
    ("msgboxstyle.abortretryignore", 2.0),
    ("MsgBoxStyle.AbortRetryIgnore", 2.0),
    ("microsoft.visualbasic.msgboxstyle.abortretryignore", 2.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.AbortRetryIgnore", 2.0),
    ("msgboxstyle.yesnocancel", 3.0),
    ("MsgBoxStyle.YesNoCancel", 3.0),
    ("microsoft.visualbasic.msgboxstyle.yesnocancel", 3.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.YesNoCancel", 3.0),
    ("msgboxstyle.yesno", 4.0),
    ("MsgBoxStyle.YesNo", 4.0),
    ("microsoft.visualbasic.msgboxstyle.yesno", 4.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.YesNo", 4.0),
    ("msgboxstyle.retrycancel", 5.0),
    ("MsgBoxStyle.RetryCancel", 5.0),
    ("microsoft.visualbasic.msgboxstyle.retrycancel", 5.0),
    ("Microsoft.VisualBasic.MsgBoxStyle.RetryCancel", 5.0),
    ("msgboxresult.ok", 1.0),
    ("MsgBoxResult.Ok", 1.0),
    ("microsoft.visualbasic.msgboxresult.ok", 1.0),
    ("Microsoft.VisualBasic.MsgBoxResult.Ok", 1.0),
    (
        "system.text.regularexpressions.regexoptions.ignorecase",
        1.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.IgnoreCase",
        1.0,
    ),
    ("regexoptions.multiline", 2.0),
    ("RegexOptions.Multiline", 2.0),
    ("system.text.regularexpressions.regexoptions.multiline", 2.0),
    ("System.Text.RegularExpressions.RegexOptions.Multiline", 2.0),
    ("regexoptions.explicitcapture", 4.0),
    ("RegexOptions.ExplicitCapture", 4.0),
    (
        "system.text.regularexpressions.regexoptions.explicitcapture",
        4.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.ExplicitCapture",
        4.0,
    ),
    ("regexoptions.compiled", 8.0),
    ("RegexOptions.Compiled", 8.0),
    ("system.text.regularexpressions.regexoptions.compiled", 8.0),
    ("System.Text.RegularExpressions.RegexOptions.Compiled", 8.0),
    ("regexoptions.singleline", 16.0),
    ("RegexOptions.Singleline", 16.0),
    (
        "system.text.regularexpressions.regexoptions.singleline",
        16.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.Singleline",
        16.0,
    ),
    ("regexoptions.ignorepatternwhitespace", 32.0),
    ("RegexOptions.IgnorePatternWhitespace", 32.0),
    (
        "system.text.regularexpressions.regexoptions.ignorepatternwhitespace",
        32.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.IgnorePatternWhitespace",
        32.0,
    ),
    ("regexoptions.righttoleft", 64.0),
    ("RegexOptions.RightToLeft", 64.0),
    (
        "system.text.regularexpressions.regexoptions.righttoleft",
        64.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.RightToLeft",
        64.0,
    ),
    ("regexoptions.ecmascript", 256.0),
    ("RegexOptions.ECMAScript", 256.0),
    (
        "system.text.regularexpressions.regexoptions.ecmascript",
        256.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.ECMAScript",
        256.0,
    ),
    ("regexoptions.cultureinvariant", 512.0),
    ("RegexOptions.CultureInvariant", 512.0),
    (
        "system.text.regularexpressions.regexoptions.cultureinvariant",
        512.0,
    ),
    (
        "System.Text.RegularExpressions.RegexOptions.CultureInvariant",
        512.0,
    ),
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

pub fn namespace_constants() -> &'static [(&'static str, f64)] {
    NAMESPACE_CONSTANTS
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
        assert!(known_type_mappings()
            .iter()
            .any(|mapping| mapping.name == "stringbuilder"));
        assert!(!known_type_mappings()
            .iter()
            .any(|mapping| mapping.name == "form"));
    }
}
