use super::helpers::{compile_ok, parse_ok};

#[test]
fn register_anonymous_module() {
    compile_ok(
        r#"
(module (func (export "f") (result i32) i32.const 42))
(register "lib")
(assert_return (invoke "lib" "f") (i32.const 42))
"#,
    ); // invoke "lib" "f" is technically not valid syntax, valid is invoke $lib "f" or just using import "lib" "f". Let's test imports
}

#[test]
fn register_and_import() {
    compile_ok(
        r#"
(module (func (export "f") (result i32) i32.const 42))
(register "lib")
(module 
  (import "lib" "f" (func $f (result i32)))
  (func (export "g") (result i32) call $f)
)
(assert_return (invoke "g") (i32.const 42))
"#,
    );
}

#[test]
fn register_named_module() {
    compile_ok(
        r#"
(module $lib (func (export "f") (result i32) i32.const 42))
(register "lib2" $lib)
(module 
  (import "lib2" "f" (func $f (result i32)))
  (func (export "g") (result i32) call $f)
)
(assert_return (invoke "g") (i32.const 42))
"#,
    );
}

#[test]
fn get_global_anonymous_module() {
    compile_ok(
        r#"
(module (global (export "g") i32 (i32.const 42)))
(assert_return (get "g") (i32.const 42))
"#,
    );
}

#[test]
fn get_global_named_module() {
    compile_ok(
        r#"
(module $m (global (export "g") i32 (i32.const 42)))
(assert_return (get $m "g") (i32.const 42))
"#,
    );
}

#[test]
fn get_global_multiple_modules() {
    compile_ok(
        r#"
(module $m1 (global (export "g") i32 (i32.const 42)))
(module $m2 (global (export "g") i32 (i32.const 99)))
(assert_return (get $m1 "g") (i32.const 42))
(assert_return (get $m2 "g") (i32.const 99))
"#,
    );
}

#[test]
fn get_global_mut() {
    compile_ok(
        r#"
(module 
  (global (export "g") (mut i32) (i32.const 42))
  (func (export "set") (param i32) local.get 0 global.set 0)
)
(assert_return (get "g") (i32.const 42))
(invoke "set" (i32.const 99))
(assert_return (get "g") (i32.const 99))
"#,
    );
}

#[test]
fn get_global_f32() {
    compile_ok(
        r#"
(module (global (export "g") f32 (f32.const 3.14)))
(assert_return (get "g") (f32.const 3.14))
"#,
    );
}

#[test]
fn get_global_f64() {
    compile_ok(
        r#"
(module (global (export "g") f64 (f64.const 2.718)))
(assert_return (get "g") (f64.const 2.718))
"#,
    );
}

#[test]
fn get_global_i64() {
    compile_ok(
        r#"
(module (global (export "g") i64 (i64.const 9999999999)))
(assert_return (get "g") (i64.const 9999999999))
"#,
    );
}
