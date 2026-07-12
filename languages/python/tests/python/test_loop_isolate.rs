//! Regression tests for the for-in + early-return bug.
//!
//! Until 2026-04-26, calling any function with an early `return` inside
//! a for-in loop crashed with "Invalid opcode 0x00 0x01" after the first
//! iteration. The VM's label_stack is global across frames; RETURN didn't
//! pop function-local BLOCK labels, so the for-in loop's `br_label 0`
//! later targeted a stale callee BLOCK and jumped into garbage bytecode.
//! Fix: compiler emits END opcodes before RETURN to drain function-local
//! labels (see `Compiler::emit_return`).

use crate::helpers::run_python;

#[test]
fn loop_norec() {
    let out = run_python(
        r#"
for i in range(3):
    print(i)
"#,
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn loop_norec_call() {
    let out = run_python(
        r#"
def g(n):
    return n + 1
for i in range(3):
    print(g(i))
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn loop_rec_call() {
    let out = run_python(
        r#"
def f(n):
    if n <= 0:
        return n
    return f(n - 1)
for i in range(3):
    print(f(i))
"#,
    );
    assert_eq!(out, vec!["0", "0", "0"]);
}

#[test]
fn loop_call_with_if_early_return() {
    let out = run_python(
        r#"
def h(n):
    if n <= 0:
        return n
    return n + 100
for i in range(3):
    print(h(i))
"#,
    );
    assert_eq!(out, vec!["0", "101", "102"]);
}

#[test]
fn loop_call_two_paths_returning() {
    let out = run_python(
        r#"
def k(n):
    if n <= 0:
        return n
    return n
for i in range(3):
    print(k(i))
"#,
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}
