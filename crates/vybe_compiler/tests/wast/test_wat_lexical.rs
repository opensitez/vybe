/// Tests for lexical analysis, tokenization, and S-Expression syntax.
use super::helpers::{compile_ok, parse_err, parse_ok};

// ── Integers & Bases ──────────────────────────────────────────────────────────

#[test]
fn int_decimal_positive() {
    parse_ok("(module (global i32 (i32.const +12345)))");
}

#[test]
fn int_decimal_negative() {
    parse_ok("(module (global i32 (i32.const -9876)))");
}

#[test]
fn int_decimal_zero() {
    parse_ok("(module (global i32 (i32.const 0)))");
    parse_ok("(module (global i32 (i32.const -0)))");
    parse_ok("(module (global i32 (i32.const +0)))");
}

#[test]
fn int_hexadecimal() {
    parse_ok("(module (global i32 (i32.const 0x1f)))");
    parse_ok("(module (global i32 (i32.const -0x1f)))");
    parse_ok("(module (global i32 (i32.const +0xabcdef)))");
}

#[test]
fn int_underscores() {
    parse_ok("(module (global i32 (i32.const 1_234_567)))");
    parse_ok("(module (global i32 (i32.const -0xab_cd_ef)))");
}

#[test]
fn int_invalid_octal_binary_wat() {
    // WAT does not support 0o or 0b prefixes for integers in core spec (must be decimal or hex)
    parse_err("(module (global i32 (i32.const 0b1010)))");
    parse_err("(module (global i32 (i32.const 0o777)))");
}

// ── Floats, Infinities, NaNs ──────────────────────────────────────────────────

#[test]
fn float_decimals() {
    parse_ok("(module (global f32 (f32.const 3.14159)))");
    parse_ok("(module (global f64 (f64.const -0.00123)))");
    parse_ok("(module (global f64 (f64.const +1.0)))");
}

#[test]
fn float_exponents() {
    parse_ok("(module (global f32 (f32.const 1e10)))");
    parse_ok("(module (global f32 (f32.const 1.2e-5)))");
    parse_ok("(module (global f64 (f64.const -3.14e+2)))");
}

#[test]
fn float_hex() {
    parse_ok("(module (global f32 (f32.const 0x1.5p+3)))");
    parse_ok("(module (global f64 (f64.const -0x0.3p-1)))");
    parse_ok("(module (global f64 (f64.const +0x1.abcde_fp+10)))");
}

#[test]
fn float_infinities() {
    parse_ok("(module (global f32 (f32.const inf)))");
    parse_ok("(module (global f32 (f32.const -inf)))");
    parse_ok("(module (global f64 (f64.const +inf)))");
}

#[test]
fn float_nans() {
    parse_ok("(module (global f32 (f32.const nan)))");
    parse_ok("(module (global f32 (f32.const -nan)))");
    parse_ok("(module (global f64 (f64.const +nan)))");
}

#[test]
fn float_nan_payloads() {
    parse_ok("(module (global f32 (f32.const nan:0x200000)))");
    parse_ok("(module (global f64 (f64.const -nan:0x1fffff)))");
}

#[test]
fn float_invalid_formats() {
    parse_err("(module (global f32 (f32.const 1.)))");
    parse_err("(module (global f32 (f32.const .5)))");
    parse_err("(module (global f32 (f32.const 0x1.5)))"); // hex float needs exponent p
}

// ── Strings & Escapes ─────────────────────────────────────────────────────────

#[test]
fn string_empty() {
    parse_ok("(module (import \"\" \"\" (func)))");
}

#[test]
fn string_utf8() {
    parse_ok("(module (import \"env\" \"hello 🌍\" (func)))");
}

#[test]
fn string_escapes_basic() {
    parse_ok("(module (import \"env\" \"line\\nfeed\" (func)))");
    parse_ok("(module (import \"env\" \"tab\\tcharacter\" (func)))");
    parse_ok("(module (import \"env\" \"quote\\\"char\" (func)))");
    parse_ok("(module (import \"env\" \"backslash\\\\char\" (func)))");
}

#[test]
fn string_hex_escapes() {
    parse_ok("(module (import \"env\" \"\\41\\42\\43\" (func)))"); // ABC
    parse_ok("(module (import \"env\" \"\\00\\ff\" (func)))");
}

#[test]
fn string_multiline() {
    parse_ok(r#"(module (import "env" "multi
line
string" (func)))"#);
}

#[test]
fn string_invalid_escapes() {
    parse_err("(module (import \"env\" \"\\z\" (func)))");
    parse_err("(module (import \"env\" \"\\f\" (func)))"); // only standard escapes + hex are valid
}

// ── Comments ──────────────────────────────────────────────────────────────────

#[test]
fn comment_line() {
    parse_ok(";; line comment\n(module)");
    parse_ok("(module ;; line comment inside\n)");
}

#[test]
fn comment_block() {
    parse_ok("(; block comment ;) (module)");
    parse_ok("(module (; block comment inside ;))");
}

#[test]
fn comment_nested_block() {
    parse_ok("(; outer (; inner ;) outer ;) (module)");
    parse_ok("(module (; multi-level (; nesting (; check ;) here ;) ;) )");
}

#[test]
fn comment_unmatched_block_err() {
    parse_err("(; unmatched block (module)");
    parse_err("(module (; nested unmatched block (; inside ;) )");
}

// ── Identifiers & Names ───────────────────────────────────────────────────────

#[test]
fn id_basic() {
    parse_ok("(module (func $func_name))");
}

#[test]
fn id_special_chars() {
    parse_ok("(module (func $func-name!@#%^&*()_+{}|:<>?-=[]\\;',./))");
}

#[test]
fn id_utf8() {
    parse_ok("(module (func $функция))");
    parse_ok("(module (func $函数))");
}

#[test]
fn id_numeric_indices() {
    parse_ok("(module (func $0) (func $1) (func (export \"test\") call $0 call $1))");
}

#[test]
fn id_invalid_identifiers() {
    parse_err("(module (func $))"); // $ alone is not an identifier
}

// ── S-Expression Balanced Checks ──────────────────────────────────────────────

#[test]
fn sexpr_empty() {
    parse_ok("()");
    parse_ok("((()))");
}

#[test]
fn sexpr_unbalanced_right() {
    parse_err("(");
    parse_err("((())");
}

#[test]
fn sexpr_unbalanced_left() {
    parse_err(")");
    parse_err("((()))(");
}
