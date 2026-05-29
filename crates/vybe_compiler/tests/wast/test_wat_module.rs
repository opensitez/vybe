/// Tests for WAT module structure — types, imports, exports, globals, memory, tables
use super::helpers::{parse_ok, compile_ok};

// ── Empty module ──────────────────────────────────────────────────────────────

#[test] fn empty_module()          { parse_ok("(module)"); }
#[test] fn named_module()          { parse_ok("(module $m)"); }
#[test] fn empty_module_compiles() { compile_ok("(module)"); }

// ── Type definitions ──────────────────────────────────────────────────────────

#[test]
fn type_def_no_params() {
    parse_ok("(module (type (func)))");
}

#[test]
fn type_def_with_params() {
    parse_ok("(module (type $add_t (func (param i32 i32) (result i32))))");
}

#[test]
fn type_def_multi_result() {
    parse_ok("(module (type (func (param i32) (result i32 i32))))");
}

#[test]
fn type_def_named_params() {
    parse_ok("(module (type (func (param $x i32) (param $y i32) (result i32))))");
}

// ── Function definitions ──────────────────────────────────────────────────────

#[test]
fn func_no_params() {
    parse_ok("(module (func))");
}

#[test]
fn func_with_params_result() {
    parse_ok("(module (func (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))");
}

#[test]
fn func_named() {
    parse_ok("(module (func $add (param $a i32) (param $b i32) (result i32) local.get $a local.get $b i32.add))");
}

#[test]
fn func_with_locals() {
    parse_ok("(module (func (param i32) (result i32) (local i32) local.get 0 local.set 1 local.get 1))");
}

#[test]
fn func_named_locals() {
    parse_ok("(module (func (local $x i32) (local $y f64)))");
}

#[test]
fn func_multiple_locals() {
    parse_ok("(module (func (local i32 i32 f64)))");
}

#[test]
fn func_with_type_ref() {
    parse_ok("(module (type $t (func (param i32) (result i32))) (func (type $t) local.get 0))");
}

#[test]
fn func_compiles() {
    compile_ok("(module (func $add (param $a i32) (param $b i32) (result i32) local.get $a local.get $b i32.add))");
}

// ── Inline exports ────────────────────────────────────────────────────────────

#[test]
fn func_inline_export() {
    parse_ok(r#"(module (func (export "add") (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))"#);
}

#[test]
fn func_multiple_inline_exports() {
    parse_ok(r#"(module (func (export "add") (export "sum") (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))"#);
}

// ── Explicit exports ──────────────────────────────────────────────────────────

#[test]
fn export_func() {
    parse_ok(r#"(module (func $f) (export "f" (func $f)))"#);
}

#[test]
fn export_func_by_index() {
    parse_ok(r#"(module (func) (export "f" (func 0)))"#);
}

#[test]
fn export_memory() {
    parse_ok(r#"(module (memory 1) (export "mem" (memory 0)))"#);
}

#[test]
fn export_global() {
    parse_ok(r#"(module (global i32 (i32.const 42)) (export "g" (global 0)))"#);
}

// ── Imports ───────────────────────────────────────────────────────────────────

#[test]
fn import_func() {
    parse_ok(r#"(module (import "env" "log" (func (param i32))))"#);
}

#[test]
fn import_func_named() {
    parse_ok(r#"(module (import "env" "log" (func $log (param i32))))"#);
}

#[test]
fn import_memory() {
    parse_ok(r#"(module (import "env" "mem" (memory 1)))"#);
}

#[test]
fn import_global() {
    parse_ok(r#"(module (import "env" "g" (global i32)))"#);
}

#[test]
fn import_global_mutable() {
    parse_ok(r#"(module (import "env" "g" (global (mut i32))))"#);
}

#[test]
fn import_table() {
    parse_ok(r#"(module (import "env" "t" (table 1 funcref)))"#);
}

#[test]
fn inline_import_func() {
    parse_ok(r#"(module (func $log (import "env" "log") (param i32)))"#);
}

// ── Globals ───────────────────────────────────────────────────────────────────

#[test]
fn global_immutable_i32() {
    parse_ok("(module (global i32 (i32.const 42)))");
}

#[test]
fn global_mutable_i32() {
    parse_ok("(module (global (mut i32) (i32.const 0)))");
}

#[test]
fn global_named() {
    parse_ok("(module (global $g i32 (i32.const 100)))");
}

#[test]
fn global_f64() {
    parse_ok("(module (global f64 (f64.const 3.14)))");
}

#[test]
fn global_compiles() {
    compile_ok("(module (global $g i32 (i32.const 42)))");
}

// ── Memory ────────────────────────────────────────────────────────────────────

#[test]
fn memory_min_only() {
    parse_ok("(module (memory 1))");
}

#[test]
fn memory_min_max() {
    parse_ok("(module (memory 1 4))");
}

#[test]
fn memory_named() {
    parse_ok("(module (memory $m 1))");
}

#[test]
fn memory_inline_export() {
    parse_ok(r#"(module (memory (export "mem") 1))"#);
}

// ── Tables ────────────────────────────────────────────────────────────────────

#[test]
fn table_funcref() {
    parse_ok("(module (table 10 funcref))");
}

#[test]
fn table_min_max() {
    parse_ok("(module (table 1 10 funcref))");
}

#[test]
fn table_named() {
    parse_ok("(module (table $t 10 funcref))");
}

// ── Data segments ─────────────────────────────────────────────────────────────

#[test]
fn data_passive() {
    parse_ok(r#"(module (data "hello"))"#);
}

#[test]
fn data_active() {
    parse_ok(r#"(module (memory 1) (data (offset (i32.const 0)) "hello"))"#);
}

#[test]
fn data_named() {
    parse_ok(r#"(module (memory 1) (data $d (offset (i32.const 0)) "world"))"#);
}

// ── Element segments ──────────────────────────────────────────────────────────

#[test]
fn elem_passive_func() {
    parse_ok("(module (func $f) (elem func $f))");
}

#[test]
fn elem_declare() {
    parse_ok("(module (func $f) (elem declare func $f))");
}

// ── Start function ────────────────────────────────────────────────────────────

#[test]
fn start_by_index() {
    parse_ok("(module (func) (start 0))");
}

#[test]
fn start_by_name() {
    parse_ok("(module (func $init) (start $init))");
}
