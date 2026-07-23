//! WASM text-format abbreviations (spec 6.4) — inline exports/imports, implicit
//! type uses, abbreviated data/elem, and folded instruction forms all expand to
//! the same module.
use super::helpers::parse_ok;
use crate::wat_exec;

wat_exec! {
    test_inline_export_on_func => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") i32.const 42 call $log))"#, "42" },
    test_folded_call_expression => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") (call $log (i32.add (i32.const 19) (i32.const 23)))))"#, "42" },
    test_folded_nested_arithmetic => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start")
          (call $log (i32.mul (i32.add (i32.const 2) (i32.const 3)) (i32.const 4)))))"#, "20" },
    test_folded_if_expression => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start")
          (call $log (if (result i32) (i32.const 1) (then (i32.const 7)) (else (i32.const 8))))))"#, "7" },
    test_implicit_type_from_params => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $add (param i32 i32) (result i32) (i32.add (local.get 0) (local.get 1)))
        (func (export "_start") (call $log (call $add (i32.const 40) (i32.const 2)))))"#, "42" },
    test_abbreviated_data_string => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory 1) (data (i32.const 0) "\2a\00\00\00")
        (func (export "_start") (call $log (i32.load (i32.const 0)))))"#, "42" },
    test_folded_local_tee => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start") (local $x i32)
          (call $log (i32.add (local.tee $x (i32.const 10)) (i32.const 5)))))"#, "15" },
    test_folded_block_with_result => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func (export "_start")
          (call $log (block (result i32) (i32.const 100)))))"#, "100" },
}

// ── Pure-syntax abbreviations that should parse ──────────────────────────────
#[test]
fn inline_import_abbreviation_parses() {
    parse_ok(r#"(module (func $f (import "m" "n") (param i32) (result i32)))"#);
}
#[test]
fn inline_memory_export_parses() {
    parse_ok(r#"(module (memory (export "mem") 1))"#);
}
#[test]
fn inline_global_export_parses() {
    parse_ok(r#"(module (global (export "g") i32 (i32.const 0)))"#);
}
#[test]
fn inline_table_export_parses() {
    parse_ok(r#"(module (table (export "t") 1 funcref))"#);
}
#[test]
fn multiple_inline_exports_parse() {
    parse_ok(r#"(module (func $f (export "a") (export "b") (result i32) i32.const 1))"#);
}
#[test]
fn abbreviated_typeuse_reference() {
    parse_ok(
        r#"(module (type $t (func (param i32) (result i32))) (func (type $t) (param i32) (result i32) local.get 0))"#,
    );
}
