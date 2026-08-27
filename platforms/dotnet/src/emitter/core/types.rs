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
    // ⛔ `Epsilon` is the smallest SUBNORMAL, not `f64::MIN_POSITIVE` (the
    // smallest normal). .NET documents 4.94065645841247E-324 for Double and
    // 1.401298E-45 for Single, and the two differ — one shared value answered
    // both wrongly.
    ("double.Epsilon", 5e-324),
    ("float.Epsilon", 1.401298464324817e-45),
    ("single.Epsilon", 1.401298464324817e-45),
    ("single.NaN", f64::NAN),
    ("single.PositiveInfinity", f64::INFINITY),
    ("single.NegativeInfinity", f64::NEG_INFINITY),
    ("single.MaxValue", 3.4028235e38),
    ("single.MinValue", -3.4028235e38),
    ("char.MaxValue", 65535.0),
    ("char.MinValue", 0.0),
    // `System.Globalization.NumberStyles` — the flag ordinals, so a `Parse`
    // overload can READ the styles it is handed instead of guessing from arity.
    // `AllowHexSpecifier` (512) is the highest flag, which is what lets the
    // parse emitters test `styles >= 512` for "this is hexadecimal".
    ("NumberStyles.None", 0.0),
    ("NumberStyles.AllowLeadingWhite", 1.0),
    ("NumberStyles.AllowTrailingWhite", 2.0),
    ("NumberStyles.AllowLeadingSign", 4.0),
    ("NumberStyles.AllowTrailingSign", 8.0),
    ("NumberStyles.AllowParentheses", 16.0),
    ("NumberStyles.AllowDecimalPoint", 32.0),
    ("NumberStyles.AllowThousands", 64.0),
    ("NumberStyles.AllowExponent", 128.0),
    ("NumberStyles.AllowCurrencySymbol", 256.0),
    ("NumberStyles.AllowHexSpecifier", 512.0),
    ("NumberStyles.Integer", 7.0),
    ("NumberStyles.Number", 111.0),
    ("NumberStyles.Float", 167.0),
    ("NumberStyles.Currency", 383.0),
    ("NumberStyles.Any", 511.0),
    ("NumberStyles.HexNumber", 515.0),
    ("numberstyles.none", 0.0),
    ("numberstyles.allowleadingwhite", 1.0),
    ("numberstyles.allowtrailingwhite", 2.0),
    ("numberstyles.allowleadingsign", 4.0),
    ("numberstyles.allowtrailingsign", 8.0),
    ("numberstyles.allowparentheses", 16.0),
    ("numberstyles.allowdecimalpoint", 32.0),
    ("numberstyles.allowthousands", 64.0),
    ("numberstyles.allowexponent", 128.0),
    ("numberstyles.allowcurrencysymbol", 256.0),
    ("numberstyles.allowhexspecifier", 512.0),
    ("numberstyles.integer", 7.0),
    ("numberstyles.number", 111.0),
    ("numberstyles.float", 167.0),
    ("numberstyles.currency", 383.0),
    ("numberstyles.any", 511.0),
    ("numberstyles.hexnumber", 515.0),
    // ⛔ `Profile::lookup_constant` LOWERCASES the key for a case-insensitive
    // language and looks the cased name up verbatim for a case-sensitive one.
    // That is what the cased/lowercase duplicate pairs further down this table
    // are for — the lowercase row serves VB, the cased row serves C#. Every
    // limit above had only the C# spelling, so `Integer.MaxValue` and
    // `Char.MaxValue` resolved to NOTHING in VB: the first rendered empty and
    // the second trapped in `charCodeAt` under `AscW`.
    //
    // The rows below are the lowercase halves, plus the VB type names (`Integer`
    // is Int32's VB alias, not a different type).
    ("int.maxvalue", 2_147_483_647.0),
    ("int.minvalue", -2_147_483_648.0),
    ("integer.maxvalue", 2_147_483_647.0),
    ("integer.minvalue", -2_147_483_648.0),
    ("short.maxvalue", 32_767.0),
    ("short.minvalue", -32_768.0),
    ("int16.maxvalue", 32_767.0),
    ("int16.minvalue", -32_768.0),
    ("int32.maxvalue", 2_147_483_647.0),
    ("int32.minvalue", -2_147_483_648.0),
    ("byte.maxvalue", 255.0),
    ("byte.minvalue", 0.0),
    ("double.maxvalue", f64::MAX),
    ("double.minvalue", -f64::MAX),
    ("double.nan", f64::NAN),
    ("double.positiveinfinity", f64::INFINITY),
    ("double.negativeinfinity", f64::NEG_INFINITY),
    ("float.maxvalue", 3.4028235e38),
    ("float.minvalue", -3.4028235e38),
    ("single.maxvalue", 3.4028235e38),
    ("single.minvalue", -3.4028235e38),
    ("double.epsilon", 5e-324),
    ("float.epsilon", 1.401298464324817e-45),
    ("single.epsilon", 1.401298464324817e-45),
    ("single.nan", f64::NAN),
    ("single.positiveinfinity", f64::INFINITY),
    ("single.negativeinfinity", f64::NEG_INFINITY),
    // The CODE UNIT, not the one-character string .NET's `Char.MaxValue` really
    // is — this table is f64-only, and every measured use is a limit comparison
    // or an `AscW`.
    ("char.maxvalue", 65535.0),
    ("char.minvalue", 0.0),
    // ⛔ `Long`/`ULong`/`Decimal` limits are deliberately ABSENT. Their maxima
    // are not representable in f64: `Int64.MaxValue` would come back as
    // ...808 rather than ...807. A missing constant is a loud failure; a
    // silently-off-by-one one is not.
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
    // `System.IO.SeekOrigin` — `Begin`/`Current`/`End`, the values
    // `MemoryStream.Seek` branches on.
    ("seekorigin.begin", 0.0),
    ("SeekOrigin.Begin", 0.0),
    ("system.io.seekorigin.begin", 0.0),
    ("System.IO.SeekOrigin.Begin", 0.0),
    ("seekorigin.current", 1.0),
    ("SeekOrigin.Current", 1.0),
    ("system.io.seekorigin.current", 1.0),
    ("System.IO.SeekOrigin.Current", 1.0),
    ("seekorigin.end", 2.0),
    ("SeekOrigin.End", 2.0),
    ("system.io.seekorigin.end", 2.0),
    ("System.IO.SeekOrigin.End", 2.0),
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
