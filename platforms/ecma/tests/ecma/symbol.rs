//! Behaviour tests for `ecma:symbol` host imports.
//!
//! Reference: ECMA-262 §20.4 Symbol.
//!
//! Each test covers a distinct behaviour.

use std::sync::Arc;
use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};
use vybe_bytecode::capabilities::Capabilities;
use vybe_compiler::compiler::platforms::register_platforms;

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-symbol-test>");
    let import_idx = chunk.add_import("ecma:symbol", name);
    let argc = args.len() as u8;
    for value in args {
        let c = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, c, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

// ── Symbol() — uniqueness guarantee ──────────────────────────────────────────

#[test]
fn symbol_constructor_returns_non_null_non_undefined() {
    let sym = invoke("Symbol", vec![]);
    assert!(!matches!(sym, Value::Null | Value::Undefined));
}

#[test]
fn two_symbol_calls_produce_distinct_values() {
    // The defining property of Symbol: every call returns a unique value.
    let a = invoke("Symbol", vec![]);
    let b = invoke("Symbol", vec![]);
    assert_ne!(a, b, "Symbol() must always produce a unique value");
}

#[test]
fn symbol_with_description_is_still_unique() {
    // Description does not affect uniqueness.
    let a = invoke("Symbol", vec![s("tag")]);
    let b = invoke("Symbol", vec![s("tag")]);
    assert_ne!(a, b);
}

// ── Symbol.for / Symbol.keyFor — global registry ──────────────────────────────

#[test]
fn symbol_for_same_key_returns_same_symbol() {
    let a = invoke("for", vec![s("shared")]);
    let b = invoke("for", vec![s("shared")]);
    assert_eq!(
        a, b,
        "Symbol.for() must return the same symbol for the same key"
    );
}

#[test]
fn symbol_for_different_keys_returns_different_symbols() {
    let a = invoke("for", vec![s("key-a")]);
    let b = invoke("for", vec![s("key-b")]);
    assert_ne!(a, b);
}

#[test]
fn key_for_returns_the_registration_key() {
    let sym = invoke("for", vec![s("my-key")]);
    assert_eq!(invoke("keyFor", vec![sym]), s("my-key"));
}

#[test]
fn key_for_on_local_symbol_returns_undefined() {
    // A symbol created with Symbol() (not Symbol.for) is not in the global registry.
    let local = invoke("Symbol", vec![s("local")]);
    assert_eq!(invoke("keyFor", vec![local]), Value::Undefined);
}

// ── Well-known symbols — existence and distinctness ───────────────────────────

#[test]
fn well_known_iterator_symbol_is_defined() {
    assert!(!matches!(
        invoke("iterator", vec![]),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn well_known_async_iterator_symbol_is_defined() {
    assert!(!matches!(
        invoke("asyncIterator", vec![]),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn well_known_to_primitive_symbol_is_defined() {
    assert!(!matches!(
        invoke("toPrimitive", vec![]),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn well_known_has_instance_symbol_is_defined() {
    assert!(!matches!(
        invoke("hasInstance", vec![]),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn well_known_to_string_tag_symbol_is_defined() {
    assert!(!matches!(
        invoke("toStringTag", vec![]),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn well_known_species_symbol_is_defined() {
    assert!(!matches!(
        invoke("species", vec![]),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn well_known_match_symbol_is_defined() {
    assert!(!matches!(
        invoke("match", vec![]),
        Value::Null | Value::Undefined
    ));
}

#[test]
fn well_known_iterator_and_async_iterator_are_distinct() {
    // Two different well-known symbols must not be the same value.
    let iter = invoke("iterator", vec![]);
    let async_iter = invoke("asyncIterator", vec![]);
    assert_ne!(iter, async_iter);
}

#[test]
fn well_known_match_and_match_all_are_distinct() {
    let m = invoke("match", vec![]);
    let ma = invoke("matchAll", vec![]);
    assert_ne!(m, ma);
}

#[test]
fn well_known_replace_search_split_are_distinct_from_each_other() {
    let replace = invoke("replace", vec![]);
    let search = invoke("search", vec![]);
    let split = invoke("split", vec![]);
    assert_ne!(replace, search);
    assert_ne!(search, split);
    assert_ne!(replace, split);
}

#[test]
fn dispose_and_async_dispose_are_distinct() {
    let d = invoke("dispose", vec![]);
    let ad = invoke("asyncDispose", vec![]);
    assert_ne!(d, ad);
}

// ── Symbol.prototype.description (§20.4.3.2) ─────────────────────────────────

#[test]
fn description_returns_the_string_passed_to_constructor() {
    // ECMA-262 §20.4.3.2: Symbol.prototype.description returns the optional description string.
    let sym = invoke("new", vec![s("my-symbol")]);
    assert_eq!(invoke("description", vec![sym]), s("my-symbol"));
}

#[test]
fn description_of_anonymous_symbol_is_undefined() {
    // Symbol() with no argument has description = undefined.
    let sym = invoke("new", vec![]);
    assert_eq!(invoke("description", vec![sym]), Value::Undefined);
}

// ── Symbol.prototype.toString (§20.4.3.3) ────────────────────────────────────

#[test]
fn to_string_wraps_description_in_symbol_parens() {
    // Symbol("foo").toString() = "Symbol(foo)".
    let sym = invoke("new", vec![s("foo")]);
    match invoke("toString", vec![sym]) {
        Value::String(s) => assert_eq!(s.as_ref(), "Symbol(foo)"),
        other => panic!("expected string, got {:?}", other),
    }
}

#[test]
fn to_string_of_anonymous_symbol_is_symbol_empty_parens() {
    // Symbol().toString() = "Symbol()".
    let sym = invoke("new", vec![]);
    match invoke("toString", vec![sym]) {
        Value::String(s) => assert_eq!(s.as_ref(), "Symbol()"),
        other => panic!("expected string, got {:?}", other),
    }
}

// ── Symbol.prototype.valueOf (§20.4.3.4) ─────────────────────────────────────

#[test]
fn value_of_returns_the_symbol_itself() {
    // valueOf must return the primitive Symbol value (same identity).
    let sym = invoke("new", vec![s("x")]);
    let val = invoke("valueOf", vec![sym.clone()]);
    assert_eq!(sym, val);
}

// ── Symbol.isConcatSpreadable / Symbol.unscopables ───────────────────────────

#[test]
fn is_concat_spreadable_symbol_is_defined() {
    // ECMA-262 §20.4.2.3: Symbol.isConcatSpreadable is a well-known Symbol.
    let sym = invoke("isConcatSpreadable", vec![]);
    assert!(matches!(
        sym,
        Value::String(_) | Value::I32(_) | Value::Object(_)
    ));
}

#[test]
fn unscopables_symbol_is_defined() {
    // ECMA-262 §20.4.2.13: Symbol.unscopables is a well-known Symbol.
    let sym = invoke("unscopables", vec![]);
    assert!(matches!(
        sym,
        Value::String(_) | Value::I32(_) | Value::Object(_)
    ));
}
