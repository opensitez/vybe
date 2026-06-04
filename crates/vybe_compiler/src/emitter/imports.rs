//! Centralized import resolution for cross-language function names.
//!
//! Each language compiler has its own `resolve_bare_import` (or equivalent)
//! that maps language-specific function names to `(module, name)` host imports.
//! Many of these mappings are identical across languages:
//!
//! - Python `print()`, Ruby `puts`, PHP `echo`, Dart `print()` all map to `("wasi:cli", "log")`
//! - Python `int()`, JS `parseInt()`, PHP `intval()`, Ruby `to_i` all map to `("ecma:number", "Number")`
//!
//! This module provides `resolve_common_import` as a single source of truth.
//! Language compilers can call it first, then fall back to language-specific
//! overrides for names that don't have a common mapping.

/// What a cross-language common name resolves to. Either a direct host
/// import or a compiler intrinsic that emits a multi-opcode composition.
///
/// Intrinsics let us preserve language-specific semantics (`cint(3.7) = 3`,
/// not 3.7) while still routing through a single resolution point. The
/// underlying intrinsic arms in `Compiler::emit_intrinsic` build the
/// composition out of ECMA primitives + WASM opcodes — no `vybe:*` host fn.
pub enum CommonImport {
    Host(&'static str, &'static str),
    Intrinsic(&'static str),
}

/// Resolve common cross-language function names to either a host import
/// or a compiler intrinsic. Returns `None` for language-specific names
/// that the caller should resolve via its own profile.
///
/// The lookup is case-insensitive to handle VB (WriteLn), PHP (ECHO), etc.
pub fn resolve_common_import(name: &str) -> Option<CommonImport> {
    match name.to_lowercase().as_str() {
        // ── I/O ──────────────────────────────────────────────────────────
        "print" | "puts" | "echo" | "display" | "writeline" | "write" => {
            Some(CommonImport::Host("wasi:cli", "log"))
        }

        "readline" | "input" | "gets" | "prompt" => {
            Some(CommonImport::Host("wasi:cli", "readLine"))
        }

        // ── Type conversion ──────────────────────────────────────────────
        // Integer conversions: `Number(x)` then floor — preserves
        // `cint(3.7) = 3` semantics every caller expects. Routed through
        // `intrinsic:cint` so the floor + Number coercion is a single
        // arm in the compiler.
        "parseint" | "cint" | "int" | "to_i" | "intval" => Some(CommonImport::Intrinsic("cint")),

        // Character/ordinal helpers used across frontend lowerings.
        "asc" | "ord" => Some(CommonImport::Intrinsic("asc")),
        "chr" | "chr$" | "chrw" | "str_from_char_code" => {
            Some(CommonImport::Intrinsic("str_from_char_code"))
        }

        // Float conversion is `Number(x)` exactly — no truncation.
        "parsefloat" | "cdbl" | "float" | "to_f" | "floatval" => {
            Some(CommonImport::Host("ecma:number", "Number"))
        }

        // Direct numeric formatting helpers used by frontend adapters
        // (notably Fortran formatted I/O) to avoid routing through a
        // JS-level sprintf polyfill.
        "tofixed" => Some(CommonImport::Host("ecma:number", "toFixed")),
        "toexponential" => Some(CommonImport::Host("ecma:number", "toExponential")),
        "toprecision" => Some(CommonImport::Host("ecma:number", "toPrecision")),

        // String coercion — §22.1.1.1 ToString.
        "tostring" | "str" | "to_s" | "strval" | "cstr" => {
            Some(CommonImport::Host("ecma:string", "String"))
        }

        "padstart" | "padleft" => Some(CommonImport::Host("ecma:string", "padStart")),
        "padend" | "padright" => Some(CommonImport::Host("ecma:string", "padEnd")),

        // JS: isNaN, isFinite — same name across languages
        "isnan" => Some(CommonImport::Host("ecma:number", "isNaN")),
        "isfinite" => Some(CommonImport::Host("ecma:number", "isFinite")),

        // ── Encoding ─────────────────────────────────────────────────────
        "btoa" | "base64_encode" => Some(CommonImport::Host("ecma:string", "btoa")),
        "atob" | "base64_decode" => Some(CommonImport::Host("ecma:string", "atob")),

        // Only the JS spelling resolves here — PHP's `urlencode` uses
        // application/x-www-form-urlencoded semantics (space → "+",
        // RFC 1738) which differ from RFC 3986 `encodeURIComponent`
        // (space → "%20"). Let each language bind its own urlencode
        // variant via the profile so the on-the-wire bytes match the
        // language's spec.
        "encodeuricomponent" => Some(CommonImport::Host("ecma:string", "encodeURIComponent")),
        "decodeuricomponent" => Some(CommonImport::Host("ecma:string", "decodeURIComponent")),

        // ── JSON ─────────────────────────────────────────────────────────
        "json_decode" => Some(CommonImport::Host("ecma:json", "parse")),
        "json_encode" => Some(CommonImport::Host("ecma:json", "stringify")),

        // ── Environment ──────────────────────────────────────────────────
        "getenv" => Some(CommonImport::Host("wasi:cli", "getEnv")),

        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn host(r: Option<CommonImport>) -> Option<(&'static str, &'static str)> {
        match r {
            Some(CommonImport::Host(m, n)) => Some((m, n)),
            _ => None,
        }
    }
    fn intrinsic(r: Option<CommonImport>) -> Option<&'static str> {
        match r {
            Some(CommonImport::Intrinsic(n)) => Some(n),
            _ => None,
        }
    }

    #[test]
    fn common_print_variants() {
        for name in &["print", "puts", "echo", "DISPLAY", "WriteLine"] {
            assert_eq!(
                host(resolve_common_import(name)),
                Some(("wasi:cli", "log")),
                "failed for {}",
                name
            );
        }
    }

    #[test]
    fn common_int_conversion_uses_intrinsic() {
        // `cint`/`parseint`/`int`/`to_i`/`intval` must floor — using
        // `Number(x)` directly (no floor) regresses VB `CInt(3.7) = 3`.
        for name in &["parseInt", "CInt", "int", "to_i", "intval"] {
            assert_eq!(
                intrinsic(resolve_common_import(name)),
                Some("cint"),
                "failed for {}",
                name
            );
        }
    }

    #[test]
    fn unknown_returns_none() {
        assert!(resolve_common_import("my_custom_func").is_none());
        assert!(resolve_common_import("setTimeout").is_none());
    }

    #[test]
    fn encoding_variants() {
        assert_eq!(
            host(resolve_common_import("btoa")),
            Some(("ecma:string", "btoa"))
        );
        assert_eq!(
            host(resolve_common_import("base64_encode")),
            Some(("ecma:string", "btoa"))
        );
        // PHP `urlencode` differs from `encodeURIComponent` (space → +
        // vs %20) — handled per-language via the profile binding.
        assert!(resolve_common_import("urlencode").is_none());
        assert_eq!(
            host(resolve_common_import("encodeuricomponent")),
            Some(("ecma:string", "encodeURIComponent"))
        );
    }
}
