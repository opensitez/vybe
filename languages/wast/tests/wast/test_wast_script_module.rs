use super::helpers::{compile_ok, parse_ok};

#[test]
fn module_anonymous() {
    compile_ok(
        r#"
(module (func (export "f") (result i32) i32.const 42))
(assert_return (invoke "f") (i32.const 42))
"#,
    );
}

#[test]
fn module_named() {
    compile_ok(
        r#"
(module $m (func (export "f") (result i32) i32.const 42))
(assert_return (invoke $m "f") (i32.const 42))
"#,
    );
}

#[test]
fn module_binary() {
    // A simple valid WASM binary module in text representation (magic + version)
    parse_ok(
        r#"
(module binary "\00asm\01\00\00\00")
"#,
    );
}

#[test]
fn module_quote() {
    parse_ok(
        r#"
(module quote "(module (func (export \"f\") (result i32) i32.const 42))")
"#,
    );
}

#[test]
fn module_multiple_sequential() {
    compile_ok(
        r#"
(module (func (export "f") (result i32) i32.const 42))
(assert_return (invoke "f") (i32.const 42))
(module (func (export "g") (result i32) i32.const 99))
(assert_return (invoke "g") (i32.const 99))
"#,
    );
}

#[test]
fn module_multiple_named() {
    compile_ok(
        r#"
(module $m1 (func (export "f") (result i32) i32.const 42))
(module $m2 (func (export "f") (result i32) i32.const 99))
(assert_return (invoke $m1 "f") (i32.const 42))
(assert_return (invoke $m2 "f") (i32.const 99))
"#,
    );
}

#[test]
fn module_export_import_chain() {
    compile_ok(
        r#"
(module $m1 (func (export "f") (result i32) i32.const 42))
(register "lib" $m1)
(module $m2 
  (import "lib" "f" (func $f (result i32)))
  (func (export "g") (result i32) call $f i32.const 1 i32.add)
)
(assert_return (invoke $m2 "g") (i32.const 43))
"#,
    );
}

#[test]
fn module_empty() {
    compile_ok(r#"(module)"#);
}

#[test]
fn module_only_types() {
    compile_ok(r#"(module (type (func)) (type (struct)) (type (array i32)))"#);
}
