use super::helpers::{compile_ok, parse_ok};

#[test]
fn invoke_no_args() {
    compile_ok(
        r#"
(module (func (export "f") nop))
(invoke "f")
"#,
    );
}

#[test]
fn invoke_with_args() {
    compile_ok(
        r#"
(module (func (export "f") (param i32 i32) nop))
(invoke "f" (i32.const 10) (i32.const 20))
"#,
    );
}

#[test]
fn invoke_on_module_alias() {
    compile_ok(
        r#"
(module $m (func (export "f") nop))
(invoke $m "f")
"#,
    );
}

#[test]
fn invoke_with_f32_args() {
    compile_ok(
        r#"
(module (func (export "f") (param f32) nop))
(invoke "f" (f32.const 1.0))
"#,
    );
}

#[test]
fn invoke_with_f64_args() {
    compile_ok(
        r#"
(module (func (export "f") (param f64) nop))
(invoke "f" (f64.const 1.0))
"#,
    );
}

#[test]
fn invoke_with_i64_args() {
    compile_ok(
        r#"
(module (func (export "f") (param i64) nop))
(invoke "f" (i64.const 9999999999))
"#,
    );
}

#[test]
fn invoke_with_nan_args() {
    compile_ok(
        r#"
(module (func (export "f") (param f32) nop))
(invoke "f" (f32.const nan))
"#,
    );
}

#[test]
fn invoke_with_infinity_args() {
    compile_ok(
        r#"
(module (func (export "f") (param f32) nop))
(invoke "f" (f32.const inf))
(invoke "f" (f32.const -inf))
"#,
    );
}

#[test]
fn invoke_with_mixed_args() {
    compile_ok(
        r#"
(module (func (export "f") (param i32 i64 f32 f64) nop))
(invoke "f" (i32.const 1) (i64.const 2) (f32.const 3.0) (f64.const 4.0))
"#,
    );
}

#[test]
fn invoke_with_ref_args() {
    compile_ok(
        r#"
(module (func (export "f") (param funcref externref) nop))
(invoke "f" (ref.null func) (ref.null extern))
"#,
    );
}

#[test]
fn invoke_fails_on_missing_export() {
    parse_ok(
        r#"
(module)
(assert_trap (invoke "f") "unknown export")
"#,
    );
}
