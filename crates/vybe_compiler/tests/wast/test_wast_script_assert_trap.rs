use super::helpers::{compile_ok, parse_ok};

#[test]
fn assert_trap_unreachable() {
    compile_ok(
        r#"
(module (func (export "f") unreachable))
(assert_trap (invoke "f") "unreachable")
"#,
    );
}

#[test]
fn assert_trap_div_zero() {
    compile_ok(
        r#"
(module (func (export "f") (result i32) i32.const 1 i32.const 0 i32.div_s))
(assert_trap (invoke "f") "integer divide by zero")
"#,
    );
}

#[test]
fn assert_trap_div_zero_u() {
    compile_ok(
        r#"
(module (func (export "f") (result i32) i32.const 1 i32.const 0 i32.div_u))
(assert_trap (invoke "f") "integer divide by zero")
"#,
    );
}

#[test]
fn assert_trap_overflow_s() {
    compile_ok(
        r#"
(module (func (export "f") (result i32) i32.const -2147483648 i32.const -1 i32.div_s))
(assert_trap (invoke "f") "integer overflow")
"#,
    );
}

#[test]
fn assert_trap_invalid_conversion_to_integer() {
    compile_ok(
        r#"
(module (func (export "f") (result i32) f32.const nan i32.trunc_f32_s))
(assert_trap (invoke "f") "invalid conversion to integer")
"#,
    );
}

#[test]
fn assert_trap_invalid_conversion_to_integer_overflow() {
    compile_ok(
        r#"
(module (func (export "f") (result i32) f32.const 3000000000.0 i32.trunc_f32_s))
(assert_trap (invoke "f") "integer overflow")
"#,
    );
}

#[test]
fn assert_trap_memory_out_of_bounds() {
    compile_ok(
        r#"
(module (memory 1) (func (export "f") (result i32) i32.const 65536 i32.load))
(assert_trap (invoke "f") "out of bounds memory access")
"#,
    );
}

#[test]
fn assert_trap_table_out_of_bounds() {
    compile_ok(
        r#"
(module (table 1 funcref) (func (export "f") (result funcref) i32.const 1 table.get 0))
(assert_trap (invoke "f") "out of bounds table access")
"#,
    );
}

#[test]
fn assert_trap_uninitialized_element() {
    compile_ok(
        r#"
(module (type $t (func)) (table 1 funcref) (func (export "f") i32.const 0 call_indirect (type $t)))
(assert_trap (invoke "f") "uninitialized element")
"#,
    );
}

#[test]
fn assert_trap_indirect_call_type_mismatch() {
    compile_ok(
        r#"
(module 
  (type $t1 (func (result i32)))
  (type $t2 (func (result f32)))
  (table 1 funcref)
  (func $g (type $t1) i32.const 0)
  (elem (i32.const 0) $g)
  (func (export "f") i32.const 0 call_indirect (type $t2) drop)
)
(assert_trap (invoke "f") "indirect call type mismatch")
"#,
    );
}

#[test]
fn assert_trap_null_pointer_dereference() {
    compile_ok(
        r#"
(module (type $S (struct (field i32))) (func (export "f") ref.null $S struct.get $S 0 drop))
(assert_trap (invoke "f") "null pointer dereference")
"#, // Or null reference
    );
}

#[test]
fn assert_trap_cast_error() {
    compile_ok(
        r#"
(module 
  (type $Base (struct (field i32)))
  (type $Sub (struct_subtype (field i32) (field i32) $Base))
  (func (export "f") 
    i32.const 0 struct.new $Base
    ref.cast $Sub drop)
)
(assert_trap (invoke "f") "cast error")
"#, // Or cast failure
    );
}

#[test]
fn assert_trap_stack_exhaustion() {
    compile_ok(
        r#"
(module (func $rec (export "f") call $rec))
(assert_trap (invoke "f") "call stack exhausted")
"#,
    );
}

#[test]
fn assert_trap_module_name() {
    compile_ok(
        r#"
(module $m (func (export "f") unreachable))
(assert_trap (invoke $m "f") "unreachable")
"#,
    );
}
