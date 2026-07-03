/// Tests for module components: type sections, imports, exports, memories, tables, and start functions.
use super::helpers::{compile_ok, parse_err, parse_ok};

// ── Type Signatures ───────────────────────────────────────────────────────────

#[test]
fn type_empty() {
    parse_ok("(module (type (func)))");
}

#[test]
fn type_params_no_result() {
    parse_ok("(module (type (func (param i32))))");
    parse_ok("(module (type (func (param i32 i64 f32 f64))))");
}

#[test]
fn type_results_no_params() {
    parse_ok("(module (type (func (result i32))))");
    parse_ok("(module (type (func (result i32 i64 f32 f64))))");
}

#[test]
fn type_params_and_results() {
    parse_ok("(module (type (func (param i32 f64) (result f32 i64))))");
}

#[test]
fn type_named_params() {
    parse_ok("(module (type (func (param $x i32) (param $y f64) (result f32))))");
}

#[test]
fn type_duplicate_definitions() {
    parse_ok("(module (type (func (param i32))) (type (func (param i32))))");
}

// ── Imports ───────────────────────────────────────────────────────────────────

#[test]
fn import_func_plain() {
    parse_ok("(module (import \"env\" \"print\" (func)))");
    parse_ok("(module (import \"env\" \"log\" (func (param i32) (result i32))))");
}

#[test]
fn import_global_immutable() {
    parse_ok("(module (import \"env\" \"g1\" (global i32)))");
    parse_ok("(module (import \"env\" \"g2\" (global f64)))");
}

#[test]
fn import_global_mutable() {
    parse_ok("(module (import \"env\" \"g1\" (global (mut i32))))");
}

#[test]
fn import_memory() {
    parse_ok("(module (import \"env\" \"mem\" (memory 1)))");
    parse_ok("(module (import \"env\" \"mem\" (memory 1 10)))");
}

#[test]
fn import_table() {
    parse_ok("(module (import \"env\" \"tbl\" (table 1 funcref)))");
    parse_ok("(module (import \"env\" \"tbl\" (table 1 5 funcref)))");
}

// ── Exports ───────────────────────────────────────────────────────────────────

#[test]
fn export_func() {
    parse_ok("(module (func $f) (export \"func\" (func $f)))");
    parse_ok("(module (func $f) (export \"func\" (func 0)))");
}

#[test]
fn export_global() {
    parse_ok("(module (global $g i32 (i32.const 0)) (export \"g\" (global $g)))");
    parse_ok("(module (global $g i32 (i32.const 0)) (export \"g\" (global 0)))");
}

#[test]
fn export_memory() {
    parse_ok("(module (memory $m 1) (export \"mem\" (memory $m)))");
    parse_ok("(module (memory $m 1) (export \"mem\" (memory 0)))");
}

#[test]
fn export_table() {
    parse_ok("(module (table $t 1 funcref) (export \"tbl\" (table $t)))");
    parse_ok("(module (table $t 1 funcref) (export \"tbl\" (table 0)))");
}

#[test]
fn export_inline() {
    parse_ok("(module (func (export \"f\")) (global (export \"g\") i32 (i32.const 0)))");
}

// ── Memory and Tables ─────────────────────────────────────────────────────────

#[test]
fn memory_decl() {
    parse_ok("(module (memory 1))");
    parse_ok("(module (memory 1 2))");
    parse_ok("(module (memory $m 10 20))");
}

#[test]
fn table_decl() {
    parse_ok("(module (table 1 funcref))");
    parse_ok("(module (table 1 10 funcref))");
    parse_ok("(module (table $t 5 20 funcref))");
}

// ── Start Function ────────────────────────────────────────────────────────────

#[test]
fn start_decl() {
    parse_ok("(module (func $s) (start $s))");
    parse_ok("(module (func) (start 0))");
}

// ── Data & Element Segments ───────────────────────────────────────────────────

#[test]
fn data_active_segment() {
    parse_ok("(module (memory 1) (data (i32.const 0) \"hello\"))");
    parse_ok("(module (memory 1) (data (i32.const 10) \"\\01\\02\\03\"))");
}

#[test]
fn data_passive_segment() {
    parse_ok("(module (data \"passive data\"))");
    parse_ok("(module (data $d \"named passive\"))");
}

#[test]
fn elem_active_segment() {
    parse_ok("(module (table 1 funcref) (func $f) (elem (i32.const 0) $f))");
    parse_ok("(module (table 1 funcref) (func $f) (elem (i32.const 0) func $f))");
}

#[test]
fn elem_passive_segment() {
    parse_ok("(module (func $f) (elem funcref (ref.func $f)))");
    parse_ok("(module (func $f) (elem $e passive funcref (ref.func $f)))");
}

#[test]
fn elem_declarative_segment() {
    parse_ok("(module (func $f) (elem declare funcref (ref.func $f)))");
}

// ── Invalid Modules (Negative Tests) ──────────────────────────────────────────

#[test]
fn invalid_import_after_func_definition() {
    parse_err("(module (func) (import \"env\" \"print\" (func)))");
}

#[test]
fn invalid_duplicate_export_names() {
    // Parser/compiler validation checks duplicate exports
    parse_err("(module (func $f1) (func $f2) (export \"f\" (func $f1)) (export \"f\" (func $f2)))");
}

#[test]
fn invalid_unresolved_function_start() {
    parse_err("(module (start $unresolved))");
}

#[test]
fn invalid_global_immutable_write() {
    // Verify parser rejects global modifications if immutable
    parse_err("(module (global $g i32 (i32.const 0)) (func global.set $g (i32.const 1)))");
}
