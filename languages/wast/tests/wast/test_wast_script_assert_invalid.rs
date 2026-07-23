use super::helpers::{compile_ok, parse_ok};

#[test]
fn assert_invalid_type_mismatch() {
    parse_ok(r#"(assert_invalid (module (func (result i32) f32.const 1.0)) "type mismatch")"#);
}

#[test]
fn assert_invalid_unknown_local() {
    parse_ok(r#"(assert_invalid (module (func (result i32) local.get 0)) "unknown local")"#);
}

#[test]
fn assert_invalid_unknown_global() {
    parse_ok(r#"(assert_invalid (module (func (result i32) global.get 0)) "unknown global")"#);
}

#[test]
fn assert_invalid_unknown_memory() {
    parse_ok(
        r#"(assert_invalid (module (func (result i32) i32.const 0 i32.load)) "unknown memory")"#,
    );
}

#[test]
fn assert_invalid_unknown_table() {
    parse_ok(
        r#"(assert_invalid (module (func (result funcref) i32.const 0 table.get 0)) "unknown table")"#,
    );
}

#[test]
fn assert_invalid_unknown_func() {
    parse_ok(r#"(assert_invalid (module (func call 1)) "unknown function")"#);
}

#[test]
fn assert_invalid_unknown_type() {
    parse_ok(r#"(assert_invalid (module (func (type 1))) "unknown type")"#);
}

#[test]
fn assert_invalid_unknown_label() {
    parse_ok(r#"(assert_invalid (module (func br 0)) "unknown label")"#);
}

#[test]
fn assert_invalid_multiple_memories() {
    parse_ok(r#"(assert_invalid (module (memory 1) (memory 1)) "multiple memories")"#); // unless multi-memory is enabled
}

#[test]
fn assert_invalid_alignment() {
    parse_ok(
        r#"(assert_invalid (module (memory 1) (func i32.const 0 i32.load align=8 drop)) "alignment must not be larger than natural")"#,
    );
}

#[test]
fn assert_invalid_constant_expression() {
    parse_ok(
        r#"(assert_invalid (module (global i32 (i32.add (i32.const 1) (i32.const 2)))) "constant expression required")"#,
    );
}

#[test]
fn assert_invalid_import_after_func() {
    parse_ok(r#"(assert_invalid (module (func) (import "a" "b" (func))) "import after function")"#);
}

#[test]
fn assert_invalid_start_function_params() {
    parse_ok(
        r#"(assert_invalid (module (func $start (param i32)) (start $start)) "start function")"#,
    );
}

#[test]
fn assert_invalid_start_function_results() {
    parse_ok(
        r#"(assert_invalid (module (func $start (result i32) i32.const 0) (start $start)) "start function")"#,
    );
}

#[test]
fn assert_invalid_duplicate_export() {
    parse_ok(
        r#"(assert_invalid (module (func (export "a")) (func (export "a"))) "duplicate export name")"#,
    );
}

#[test]
fn assert_invalid_duplicate_local() {
    parse_ok(r#"(assert_invalid (module (func (local $a i32) (local $a i32))) "duplicate local")"#);
}
