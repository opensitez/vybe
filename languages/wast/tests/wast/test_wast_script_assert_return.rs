use super::helpers::{compile_ok, parse_ok};

#[test]
fn assert_return_basic_i32() {
    compile_ok(
        r#"
(module (func (export "f") (result i32) i32.const 42))
(assert_return (invoke "f") (i32.const 42))
"#,
    );
}

#[test]
fn assert_return_basic_i64() {
    compile_ok(
        r#"
(module (func (export "f") (result i64) i64.const 9999999999))
(assert_return (invoke "f") (i64.const 9999999999))
"#,
    );
}

#[test]
fn assert_return_basic_f32() {
    compile_ok(
        r#"
(module (func (export "f") (result f32) f32.const 3.14))
(assert_return (invoke "f") (f32.const 3.14))
"#,
    );
}

#[test]
fn assert_return_basic_f64() {
    compile_ok(
        r#"
(module (func (export "f") (result f64) f64.const 2.718))
(assert_return (invoke "f") (f64.const 2.718))
"#,
    );
}

#[test]
fn assert_return_multiple_results() {
    compile_ok(
        r#"
(module (func (export "f") (result i32 i32 i32) i32.const 1 i32.const 2 i32.const 3))
(assert_return (invoke "f") (i32.const 1) (i32.const 2) (i32.const 3))
"#,
    );
}

#[test]
fn assert_return_nan_canonical_f32() {
    compile_ok(
        r#"
(module (func (export "f") (result f32) f32.const nan))
(assert_return (invoke "f") (f32.const nan:canonical))
"#,
    );
}

#[test]
fn assert_return_nan_arithmetic_f32() {
    compile_ok(
        r#"
(module (func (export "f") (result f32) f32.const nan:arithmetic))
(assert_return (invoke "f") (f32.const nan:arithmetic))
"#,
    );
}

#[test]
fn assert_return_nan_canonical_f64() {
    compile_ok(
        r#"
(module (func (export "f") (result f64) f64.const nan))
(assert_return (invoke "f") (f64.const nan:canonical))
"#,
    );
}

#[test]
fn assert_return_nan_arithmetic_f64() {
    compile_ok(
        r#"
(module (func (export "f") (result f64) f64.const nan:arithmetic))
(assert_return (invoke "f") (f64.const nan:arithmetic))
"#,
    );
}

#[test]
fn assert_return_ref_null_funcref() {
    compile_ok(
        r#"
(module (func (export "f") (result funcref) ref.null func))
(assert_return (invoke "f") (ref.null func))
"#,
    );
}

#[test]
fn assert_return_ref_null_externref() {
    compile_ok(
        r#"
(module (func (export "f") (result externref) ref.null extern))
(assert_return (invoke "f") (ref.null extern))
"#,
    );
}

#[test]
fn assert_return_ref_null_anyref() {
    compile_ok(
        r#"
(module (func (export "f") (result anyref) ref.null any))
(assert_return (invoke "f") (ref.null any))
"#,
    );
}

#[test]
fn assert_return_ref_func() {
    compile_ok(
        r#"
(module 
  (func $dummy)
  (func (export "f") (result funcref) ref.func $dummy)
)
(assert_return (invoke "f") (ref.func))
"#,
    );
}

#[test]
fn assert_return_ref_extern() {
    compile_ok(
        r#"
(module 
  (func (export "f") (param externref) (result externref) local.get 0)
)
(assert_return (invoke "f") (ref.extern 1))
"#, // using ref.extern for externref tests if supported by parse
    );
}

#[test]
fn assert_return_simd_v128() {
    compile_ok(
        r#"
(module 
  (func (export "f") (result v128) v128.const i32x4 1 2 3 4)
)
(assert_return (invoke "f") (v128.const i32x4 1 2 3 4))
"#,
    );
}

#[test]
fn assert_return_args_and_results() {
    compile_ok(
        r#"
(module 
  (func (export "f") (param i32 i32) (result i32 i32) 
    local.get 1 
    local.get 0)
)
(assert_return (invoke "f" (i32.const 10) (i32.const 20)) (i32.const 20) (i32.const 10))
"#,
    );
}

#[test]
fn assert_return_module_name() {
    compile_ok(
        r#"
(module $m (func (export "f") (result i32) i32.const 42))
(assert_return (invoke $m "f") (i32.const 42))
"#,
    );
}

#[test]
fn assert_return_empty() {
    compile_ok(
        r#"
(module (func (export "f")))
(assert_return (invoke "f"))
"#,
    );
}

#[test]
fn assert_return_multiple_invokes() {
    compile_ok(
        r#"
(module 
  (func (export "f") (result i32) i32.const 1)
  (func (export "g") (result i32) i32.const 2)
)
(assert_return (invoke "f") (i32.const 1))
(assert_return (invoke "g") (i32.const 2))
"#,
    );
}
