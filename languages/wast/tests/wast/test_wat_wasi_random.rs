//! WASI random — importing get-random-bytes / get-random-u64. Values are
//! non-deterministic, so these validate the import surface parses and links.
use super::helpers::{compile_ok, parse_ok};

#[test]
fn import_get_random_u64() {
    parse_ok(r#"(module (import "wasi:random/random" "get-random-u64" (func $r (result i64))))"#);
}
#[test]
fn import_get_random_bytes() {
    parse_ok(
        r#"(module (import "wasi:random/random" "get-random-bytes" (func $r (param i64) (result i32))))"#,
    );
}
#[test]
fn import_insecure_random() {
    parse_ok(
        r#"(module (import "wasi:random/insecure" "get-insecure-random-u64" (func $r (result i64))))"#,
    );
}
#[test]
fn call_random_u64_and_drop() {
    compile_ok(
        r#"(module
          (import "wasi:random/random" "get-random-u64" (func $r (result i64)))
          (func (export "_start") call $r drop))"#,
    );
}
