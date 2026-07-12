/// Tests for WAST/WAT proposal extensions:
/// exceptions, tail calls, reference types, SIMD, bulk memory, multi-value
use super::helpers::parse_ok;

// ── Exceptions proposal ───────────────────────────────────────────────────────

#[test]
fn tag_definition() {
    parse_ok("(module (tag $e (param i32)))");
}

#[test]
fn tag_no_params() {
    parse_ok("(module (tag $e))");
}

#[test]
fn tag_export() {
    parse_ok(r#"(module (tag (export "e") (param i32)))"#);
}

#[test]
fn tag_import() {
    parse_ok(r#"(module (import "env" "e" (tag (param i32))))"#);
}

#[test]
fn throw_with_tag() {
    parse_ok(
        r#"
(module
  (tag $e (param i32))
  (func (export "f") i32.const 42 throw $e))
"#,
    );
}

#[test]
fn try_catch_plain() {
    parse_ok(
        r#"
(module
  (tag $e (param i32))
  (func (export "f") (result i32)
    try (result i32)
      i32.const 1
    catch $e
      ;; e value is on stack
    end))
"#,
    );
}

#[test]
fn try_catch_all() {
    parse_ok(
        r#"
(module
  (func (export "f")
    try
      nop
    catch_all
      nop
    end))
"#,
    );
}

#[test]
fn try_delegate() {
    parse_ok(
        r#"
(module
  (tag $e)
  (func (export "f")
    block $b
      try
        nop
      delegate $b
    end))
"#,
    );
}

#[test]
fn rethrow_in_catch() {
    parse_ok(
        r#"
(module
  (tag $e)
  (func (export "f")
    try
      nop
    catch $e
      rethrow 0
    end))
"#,
    );
}

// ── Tail calls ────────────────────────────────────────────────────────────────

#[test]
fn return_call_direct() {
    parse_ok(
        r#"
(module
  (func $f (param i32) (result i32) local.get 0)
  (func (export "g") (param i32) (result i32) local.get 0 return_call $f))
"#,
    );
}

#[test]
fn return_call_indirect_basic() {
    parse_ok(
        r#"
(module
  (type $t (func (param i32) (result i32)))
  (table 1 funcref)
  (func (export "f") (param i32 i32) (result i32)
    local.get 0 local.get 1 return_call_indirect (type $t)))
"#,
    );
}

// ── Reference types ───────────────────────────────────────────────────────────

#[test]
fn ref_null_funcref_instr() {
    parse_ok("(module (func (result funcref) ref.null funcref))");
}

#[test]
fn ref_null_externref_instr() {
    parse_ok("(module (func (result externref) ref.null externref))");
}

#[test]
fn ref_is_null_instr() {
    parse_ok("(module (func (param funcref) (result i32) local.get 0 ref.is_null))");
}

#[test]
fn ref_func_instr() {
    parse_ok("(module (func $f) (func (result funcref) ref.func $f))");
}

#[test]
fn table_funcref_type() {
    parse_ok("(module (table 10 funcref))");
}

#[test]
fn table_externref_type() {
    parse_ok("(module (table 10 externref))");
}

#[test]
fn table_get_set() {
    parse_ok(
        r#"
(module
  (table $t 10 funcref)
  (func (param i32) (result funcref) local.get 0 table.get $t)
  (func (param i32 funcref) local.get 0 local.get 1 table.set $t))
"#,
    );
}

#[test]
fn table_grow_size_fill() {
    parse_ok(
        r#"
(module
  (table $t 1 funcref)
  (func (result i32) ref.null funcref i32.const 1 table.grow $t)
  (func (result i32) table.size $t))
"#,
    );
}

// ── Bulk memory ───────────────────────────────────────────────────────────────

#[test]
fn memory_copy_instr() {
    parse_ok(
        "(module (memory 1) (func (param i32 i32 i32) local.get 0 local.get 1 local.get 2 memory.copy))",
    );
}

#[test]
fn memory_fill_instr() {
    parse_ok(
        "(module (memory 1) (func (param i32 i32 i32) local.get 0 local.get 1 local.get 2 memory.fill))",
    );
}

#[test]
fn memory_init_instr() {
    parse_ok(
        r#"(module (memory 1) (data $d "hello") (func (param i32 i32 i32) local.get 0 local.get 1 local.get 2 memory.init $d))"#,
    );
}

#[test]
fn table_copy_instr() {
    parse_ok(
        "(module (table 10 funcref) (func (param i32 i32 i32) local.get 0 local.get 1 local.get 2 table.copy))",
    );
}

#[test]
fn table_init_instr() {
    parse_ok(
        "(module (table 10 funcref) (elem $e func) (func (param i32 i32 i32) local.get 0 local.get 1 local.get 2 table.init $e))",
    );
}

// ── Multi-value ───────────────────────────────────────────────────────────────

#[test]
fn multi_value_result() {
    parse_ok("(module (func (result i32 i32) i32.const 1 i32.const 2))");
}

#[test]
fn multi_value_param_result() {
    parse_ok("(module (func (param i32 i32) (result i32 i32) local.get 0 local.get 1))");
}

#[test]
fn multi_value_type_def() {
    parse_ok("(module (type $t (func (param i32) (result i32 i32))))");
}

#[test]
fn block_multi_value() {
    parse_ok("(module (func (result i32 i32) (block (result i32 i32) i32.const 1 i32.const 2)))");
}

// ── SIMD v128 ─────────────────────────────────────────────────────────────────

#[test]
fn v128_const_i32x4() {
    parse_ok("(module (func (result v128) v128.const i32x4 1 2 3 4))");
}

#[test]
fn v128_const_f64x2() {
    parse_ok("(module (func (result v128) v128.const f64x2 1.0 2.0))");
}

#[test]
fn i32x4_splat() {
    parse_ok("(module (func (param i32) (result v128) local.get 0 i32x4.splat))");
}

#[test]
fn f32x4_add() {
    parse_ok("(module (func (param v128 v128) (result v128) local.get 0 local.get 1 f32x4.add))");
}

#[test]
fn v128_and() {
    parse_ok("(module (func (param v128 v128) (result v128) local.get 0 local.get 1 v128.and))");
}

#[test]
fn v128_not() {
    parse_ok("(module (func (param v128) (result v128) local.get 0 v128.not))");
}

#[test]
fn v128_any_true() {
    parse_ok("(module (func (param v128) (result i32) local.get 0 v128.any_true))");
}

#[test]
fn i8x16_shuffle() {
    parse_ok(
        "(module (func (param v128 v128) (result v128) local.get 0 local.get 1 i8x16.shuffle 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15))",
    );
}

// ── Named modules in WAST ─────────────────────────────────────────────────────

#[test]
fn named_module_invoke() {
    parse_ok(
        r#"
(module $m1 (func (export "f") (result i32) i32.const 1))
(module $m2 (func (export "f") (result i32) i32.const 2))
(assert_return (invoke $m1 "f") (i32.const 1))
(assert_return (invoke $m2 "f") (i32.const 2))
"#,
    );
}

#[test]
fn register_and_import() {
    parse_ok(
        r#"
(module $lib
  (func (export "add") (param i32 i32) (result i32) local.get 0 local.get 1 i32.add))
(register "lib" $lib)
(module
  (import "lib" "add" (func $add (param i32 i32) (result i32)))
  (func (export "double_add") (param i32 i32) (result i32)
    local.get 0 local.get 1 call $add
    local.get 0 local.get 1 call $add
    i32.add))
"#,
    );
}

// ── assert_suspension (threads proposal) ─────────────────────────────────────

#[test]
fn assert_suspension_parses() {
    parse_ok(
        r#"
(module (func (export "f")))
(assert_suspension (invoke "f") "suspended")
"#,
    );
}
