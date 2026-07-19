//! Breadth coverage of distinct WAST *syntactic forms* — one behavioral test
//! per construct, so each surface form of the text grammar is exercised end to
//! end (parse → compile → run), not merely parsed.
use crate::wat_exec;

wat_exec! {
    // Inline `(export …)` abbreviation on a memory.
    test_form_inline_export_memory => { r#"
(memory (export "mem") 1)
(func (export "_start")
  i32.const 0 i32.const 42 i32.store
  i32.const 0 i32.load call $log)
"#, "42" },

    // Mutable global: `(global $g (mut i32) …)` with global.set/get.
    test_form_mutable_global => { r#"
(global $g (mut i32) (i32.const 5))
(func (export "_start")
  i32.const 10 global.set $g
  global.get $g call $log)
"#, "10" },

    // Immutable global with inline export, read by name.
    test_form_global_inline_export => { r#"
(global $g (export "g") i32 (i32.const 7))
(func (export "_start") global.get $g call $log)
"#, "7" },

    // Global read/written by NUMERIC index (no `$name`).
    test_form_global_numeric_index => { r#"
(global (mut i32) (i32.const 3))
(func (export "_start")
  i32.const 8 global.set 0
  global.get 0 call $log)
"#, "8" },

    // `(start $f)` runs at instantiation; here it seeds a global _start reads.
    test_form_start_function => { r#"
(global $g (mut i32) (i32.const 0))
(func $init i32.const 99 global.set $g)
(start $init)
(func (export "_start") global.get $g call $log)
"#, "99" },

    // Table declared with the inline `(elem …)` abbreviation (implicit size +
    // active segment).
    test_form_table_inline_elem => { r#"
(type $v (func (result i32)))
(func $f (result i32) i32.const 55)
(table funcref (elem $f))
(func (export "_start") i32.const 0 call_indirect (type $v) call $log)
"#, "55" },

    // Named params AND named locals.
    test_form_named_params_locals => { r#"
(func $sq (param $x i32) (result i32) (local $y i32)
  local.get $x local.get $x i32.mul local.set $y
  local.get $y)
(func (export "_start") i32.const 6 call $sq call $log)
"#, "36" },

    // Fully folded S-expression instruction form.
    test_form_folded_expression => { r#"
(func (export "_start")
  (call $log (i32.add (i32.const 20) (i32.const 22))))
"#, "42" },

    // `call_indirect (type $sig)` through a table + active elem segment.
    test_form_call_indirect_typeuse => { r#"
(type $ii (func (param i32) (result i32)))
(func $double (param i32) (result i32) local.get 0 i32.const 2 i32.mul)
(table 1 funcref)
(elem (i32.const 0) $double)
(func (export "_start")
  i32.const 21 i32.const 0 call_indirect (type $ii) call $log)
"#, "42" },

    // `block (result i32)` typed result.
    test_form_block_result_type => { r#"
(func (export "_start")
  (block (result i32) i32.const 8) call $log)
"#, "8" },

    // Folded `if (result i32) (then …) (else …)`.
    test_form_folded_if_result => { r#"
(func (export "_start")
  (if (result i32) (i32.const 1) (then (i32.const 1)) (else (i32.const 0)))
  call $log)
"#, "1" },

    // Typed `select (result i32)` (picks the 2nd operand when cond = 0).
    test_form_typed_select => { r#"
(func (export "_start")
  i32.const 11 i32.const 22 i32.const 0 select (result i32) call $log)
"#, "22" },

    // Hex integer literal with `_` digit separators.
    test_form_hex_underscore_literal => { r#"
(func (export "_start") i32.const 0x_ff call $log)
"#, "255" },

    // Active data segment written to memory, read back as a byte.
    test_form_data_active_segment => { r#"
(memory 1)
(data (i32.const 0) "A")
(func (export "_start") i32.const 0 i32.load8_u call $log)
"#, "65" },

    // Custom annotation `(@name …)` — parsed and ignored (no runtime effect).
    test_form_custom_annotation => { r#"
(@custom "meta" "\00")
(func (export "_start") i32.const 42 call $log)
"#, "42" },

    // Structured loop with `br_if`/`br` to named labels.
    test_form_loop_labeled_branch => { r#"
(func (export "_start") (local $i i32) (local $s i32)
  (block $b (loop $l
    local.get $i i32.const 5 i32.eq br_if $b
    local.get $s local.get $i i32.add local.set $s
    local.get $i i32.const 1 i32.add local.set $i
    br $l))
  local.get $s call $log)
"#, "10" },

    // f64 float literal + f64 logging sink.
    test_form_f64_literal => { r#"
(func (export "_start") f64.const 3.14 call $log_f64)
"#, "3.14" },
}
