//! WASI clocks — importing and calling the monotonic / wall clocks. Clock
//! values are non-deterministic, so these check the import surface parses and
//! links (compiles) rather than asserting an exact timestamp.
use super::helpers::{compile_ok, parse_ok};

#[test]
fn import_monotonic_now() {
    parse_ok(
        r#"(module (import "wasi:clocks/monotonic-clock" "now" (func $now (result i64))))"#,
    );
}
#[test]
fn import_monotonic_resolution() {
    parse_ok(
        r#"(module (import "wasi:clocks/monotonic-clock" "resolution" (func $res (result i64))))"#,
    );
}
#[test]
fn import_wall_clock_now() {
    parse_ok(
        r#"(module (import "wasi:clocks/wall-clock" "now" (func $now (result i64 i32))))"#,
    );
}
#[test]
fn call_monotonic_now_and_drop() {
    compile_ok(
        r#"(module
          (import "wasi:clocks/monotonic-clock" "now" (func $now (result i64)))
          (func (export "_start") call $now drop))"#,
    );
}
#[test]
fn monotonic_now_difference_pattern() {
    compile_ok(
        r#"(module
          (import "wasi:clocks/monotonic-clock" "now" (func $now (result i64)))
          (func (export "_start") (result i64) call $now call $now i64.sub))"#,
    );
}
