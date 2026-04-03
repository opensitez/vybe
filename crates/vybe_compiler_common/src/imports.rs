//! Centralized import resolution for cross-language function names.
//!
//! Each language compiler has its own `resolve_bare_import` (or equivalent)
//! that maps language-specific function names to `(module, name)` host imports.
//! Many of these mappings are identical across languages:
//!
//! - Python `print()`, Ruby `puts`, PHP `echo`, Dart `print()` all map to `("wasi:cli", "log")`
//! - Python `int()`, JS `parseInt()`, PHP `intval()`, Ruby `to_i` all map to `("vybe:convert", "cint")`
//!
//! This module provides `resolve_common_import` as a single source of truth.
//! Language compilers can call it first, then fall back to language-specific
//! overrides for names that don't have a common mapping.

/// Resolve common cross-language function names to host module imports.
/// Returns `(module, name)` if the function maps to a known host import,
/// `None` if it's language-specific or should use a different mechanism.
///
/// The lookup is case-insensitive to handle VB (WriteLn), PHP (ECHO), etc.
pub fn resolve_common_import(name: &str) -> Option<(&'static str, &'static str)> {
    match name.to_lowercase().as_str() {
        // ── I/O ──────────────────────────────────────────────────────────
        // Python: print, Ruby: puts/print/p, PHP: echo/print, Dart: print,
        // VB: Console.WriteLine, COBOL: DISPLAY
        "print" | "puts" | "echo" | "display" | "writeline" | "write"
            => Some(("wasi:cli", "log")),

        // Python: input, Ruby: gets/readline, PHP: readline, JS: prompt
        "readline" | "input" | "gets" | "prompt"
            => Some(("wasi:cli", "readLine")),

        // ── Type conversion ──────────────────────────────────────────────
        // JS: parseInt, Python: int, Ruby: to_i, PHP: intval, VB: CInt, Dart: int.parse
        "parseint" | "cint" | "int" | "to_i" | "intval"
            => Some(("vybe:convert", "cint")),

        // JS: parseFloat, Python: float, Ruby: to_f, PHP: floatval, VB: CDbl
        "parsefloat" | "cdbl" | "float" | "to_f" | "floatval"
            => Some(("vybe:convert", "cdbl")),

        // Python: str, Ruby: to_s, PHP: strval, VB: CStr
        "tostring" | "str" | "to_s" | "strval" | "cstr"
            => Some(("vybe:convert", "toString")),

        // JS: isNaN, isFinite — same name across languages
        "isnan"    => Some(("vybe:convert", "isNaN")),
        "isfinite" => Some(("vybe:convert", "isFinite")),

        // ── Encoding ─────────────────────────────────────────────────────
        // JS: btoa/atob, PHP: base64_encode/base64_decode, Ruby: Base64.encode64
        "btoa" | "base64_encode" => Some(("vybe:convert", "btoa")),
        "atob" | "base64_decode" => Some(("vybe:convert", "atob")),

        // JS: encodeURIComponent, PHP: urlencode/rawurlencode
        "encodeuricomponent" | "urlencode" | "rawurlencode"
            => Some(("vybe:convert", "encodeURIComponent")),
        "decodeuricomponent" | "urldecode" | "rawurldecode"
            => Some(("vybe:convert", "decodeURIComponent")),

        // ── JSON ─────────────────────────────────────────────────────────
        // JS: JSON.parse, PHP: json_decode, Python: json.loads, Ruby: JSON.parse
        "json_decode" => Some(("vybe:json", "parse")),
        "json_encode" => Some(("vybe:json", "stringify")),

        // ── Environment ──────────────────────────────────────────────────
        // PHP: getenv, Python: os.getenv, Ruby: ENV[]
        "getenv" => Some(("wasi:cli", "getEnv")),

        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn common_print_variants() {
        for name in &["print", "puts", "echo", "DISPLAY", "WriteLine"] {
            let result = resolve_common_import(name);
            assert_eq!(result, Some(("wasi:cli", "log")), "failed for {}", name);
        }
    }

    #[test]
    fn common_int_conversion() {
        for name in &["parseInt", "CInt", "int", "to_i", "intval"] {
            let result = resolve_common_import(name);
            assert_eq!(result, Some(("vybe:convert", "cint")), "failed for {}", name);
        }
    }

    #[test]
    fn unknown_returns_none() {
        assert_eq!(resolve_common_import("my_custom_func"), None);
        assert_eq!(resolve_common_import("setTimeout"), None);
    }

    #[test]
    fn encoding_variants() {
        assert_eq!(resolve_common_import("btoa"), Some(("vybe:convert", "btoa")));
        assert_eq!(resolve_common_import("base64_encode"), Some(("vybe:convert", "btoa")));
        assert_eq!(resolve_common_import("urlencode"), Some(("vybe:convert", "encodeURIComponent")));
    }
}
