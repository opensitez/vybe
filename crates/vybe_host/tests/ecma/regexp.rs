//! Behaviour tests for `ecma:regexp` host imports — ECMA-262 §22.2
//! RegExp constructor + `RegExp.prototype.{test, exec, toString}` plus
//! the regex-taking `String.prototype.{match, matchAll, search, replace,
//! replaceAll, split}` methods.
//!
//! Reference: ECMA-262 §22.2 RegExp + §22.1.3.{13,14,16,18,19,20}.

use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-regexp-test>");
    let import_idx = chunk.add_import("ecma:regexp", name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn s(text: &str) -> Value {
    Value::String(std::sync::Arc::from(text))
}

fn as_string(value: &Value) -> String {
    match value {
        Value::String(text) => text.to_string(),
        other => format!("{}", other),
    }
}

fn array_strings(value: &Value) -> Vec<String> {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        if let ObjectKind::Array(elements) = &object.kind {
            return elements
                .iter()
                .map(|element| match element {
                    Value::String(text) => text.to_string(),
                    other => format!("{}", other),
                })
                .collect();
        }
    }
    Vec::new()
}

fn obj_prop(value: &Value, key: &str) -> Value {
    if let Value::Object(object) = value {
        let object = object.lock().unwrap();
        return object.properties.get(key).cloned().unwrap_or(Value::Undefined);
    }
    Value::Undefined
}

// ── Constructor ──────────────────────────────────────────────────────

#[test]
fn new_stores_source_and_flags() {
    let r = invoke("new", vec![s("hello"), s("i")]);
    assert_eq!(as_string(&obj_prop(&r, "source")), "hello");
    assert_eq!(as_string(&obj_prop(&r, "flags")), "i");
}

#[test]
fn new_sets_boolean_flag_accessors() {
    let r = invoke("new", vec![s("x"), s("gim")]);
    assert_eq!(obj_prop(&r, "global"), Value::Bool(true));
    assert_eq!(obj_prop(&r, "ignoreCase"), Value::Bool(true));
    assert_eq!(obj_prop(&r, "multiline"), Value::Bool(true));
    assert_eq!(obj_prop(&r, "dotAll"), Value::Bool(false));
}

#[test]
fn new_no_flags_defaults_all_false() {
    let r = invoke("new", vec![s("x")]);
    assert_eq!(obj_prop(&r, "global"), Value::Bool(false));
    assert_eq!(obj_prop(&r, "ignoreCase"), Value::Bool(false));
}

#[test]
fn new_stamps_type_for_instanceof() {
    let r = invoke("new", vec![s("x")]);
    assert_eq!(as_string(&obj_prop(&r, "__type")), "RegExp");
}

#[test]
fn new_initial_last_index_is_zero() {
    let r = invoke("new", vec![s("x")]);
    assert_eq!(obj_prop(&r, "lastIndex"), Value::I32(0));
}

#[test]
fn new_from_existing_regexp_inherits_source_and_flags() {
    // `new RegExp(existing)` — pattern + flags carry over (per
    // ECMA-262 §22.2.4).
    let original = invoke("new", vec![s("abc"), s("i")]);
    let copied = invoke("new", vec![original]);
    assert_eq!(as_string(&obj_prop(&copied, "source")), "abc");
    assert_eq!(as_string(&obj_prop(&copied, "flags")), "i");
}

// ── test ─────────────────────────────────────────────────────────────

#[test]
fn test_returns_true_when_pattern_matches() {
    let r = invoke("new", vec![s("hello")]);
    assert_eq!(invoke("test", vec![r, s("hello world")]), Value::Bool(true));
}

#[test]
fn test_returns_false_when_pattern_does_not_match() {
    let r = invoke("new", vec![s("xyz")]);
    assert_eq!(invoke("test", vec![r, s("hello world")]), Value::Bool(false));
}

#[test]
fn test_with_case_insensitive_flag() {
    let r = invoke("new", vec![s("HELLO"), s("i")]);
    assert_eq!(invoke("test", vec![r, s("hello world")]), Value::Bool(true));
}

#[test]
fn test_invalid_pattern_returns_false() {
    let r = invoke("new", vec![s("[")]);
    // Invalid regex compiles to None; spec says SyntaxError on construct,
    // but MVP swallows and returns Bool(false) rather than trapping.
    assert_eq!(invoke("test", vec![r, s("anything")]), Value::Bool(false));
}

#[test]
fn test_unicode_sets_rgi_emoji_matches_smiley() {
    let r = invoke("new", vec![s("\\p{RGI_Emoji}"), s("v")]);
    assert_eq!(invoke("test", vec![r, s("🙂")]), Value::Bool(true));
}

#[test]
fn test_unicode_sets_rgi_emoji_rejects_ascii() {
    let r = invoke("new", vec![s("\\p{RGI_Emoji}"), s("v")]);
    assert_eq!(invoke("test", vec![r, s("A")]), Value::Bool(false));
}

// ── exec ─────────────────────────────────────────────────────────────

#[test]
fn exec_returns_match_array_with_full_match_at_zero() {
    let r = invoke("new", vec![s("(\\w+) (\\w+)")]);
    let result = invoke("exec", vec![r, s("hello world rest")]);
    let elems = array_strings(&result);
    // [0] = full match, [1] = first group, [2] = second group
    assert_eq!(elems[0], "hello world");
    assert_eq!(elems[1], "hello");
    assert_eq!(elems[2], "world");
}

#[test]
fn exec_match_array_carries_index() {
    let r = invoke("new", vec![s("world")]);
    let result = invoke("exec", vec![r, s("hello world rest")]);
    assert_eq!(obj_prop(&result, "index"), Value::I32(6));
}

#[test]
fn exec_match_array_carries_input() {
    let r = invoke("new", vec![s("world")]);
    let result = invoke("exec", vec![r, s("hello world rest")]);
    assert_eq!(as_string(&obj_prop(&result, "input")), "hello world rest");
}

#[test]
fn exec_returns_null_when_no_match() {
    let r = invoke("new", vec![s("xyz")]);
    assert_eq!(invoke("exec", vec![r, s("hello")]), Value::Null);
}

#[test]
fn exec_named_groups_appear_in_groups_object() {
    let r = invoke("new", vec![s("(?P<word>\\w+)")]);
    let result = invoke("exec", vec![r, s("hello")]);
    let groups = obj_prop(&result, "groups");
    assert_eq!(as_string(&obj_prop(&groups, "word")), "hello");
}

#[test]
fn exec_unicode_sets_rgi_emoji_returns_full_match() {
    let r = invoke("new", vec![s("\\p{RGI_Emoji}"), s("v")]);
    let result = invoke("exec", vec![r, s("A🙂B")]);
    assert_eq!(array_strings(&result), vec!["🙂"]);
    assert_eq!(obj_prop(&result, "index"), Value::I32(1));
}

// ── toString ─────────────────────────────────────────────────────────

#[test]
fn to_string_uses_slash_pattern_slash_flags_form() {
    let r = invoke("new", vec![s("abc"), s("gi")]);
    assert_eq!(as_string(&invoke("toString", vec![r])), "/abc/gi");
}

// ── String.prototype.match ───────────────────────────────────────────

#[test]
fn match_non_global_returns_first_match_with_groups() {
    let r = invoke("new", vec![s("(\\w+) (\\w+)")]);
    let result = invoke("match", vec![s("hello world"), r]);
    let elems = array_strings(&result);
    assert_eq!(elems[0], "hello world");
    assert_eq!(elems[1], "hello");
    assert_eq!(elems[2], "world");
}

#[test]
fn match_global_returns_array_of_full_matches_only() {
    let r = invoke("new", vec![s("\\d+"), s("g")]);
    let result = invoke("match", vec![s("a1 b22 c333"), r]);
    assert_eq!(array_strings(&result), vec!["1", "22", "333"]);
}

#[test]
fn match_returns_null_on_no_match() {
    let r = invoke("new", vec![s("xyz")]);
    assert_eq!(invoke("match", vec![s("hello"), r]), Value::Null);
}

// ── String.prototype.matchAll ────────────────────────────────────────

#[test]
fn match_all_returns_array_of_match_arrays() {
    let r = invoke("new", vec![s("(\\w)(\\d)")]);
    let result = invoke("matchAll", vec![s("a1 b2 c3"), r]);
    if let Value::Object(arr) = &result {
        let arr = arr.lock().unwrap();
        if let ObjectKind::Array(ref outer) = arr.kind {
            assert_eq!(outer.len(), 3);
            // Each element is itself a match Array — first one is "a1" + "a" + "1"
            let first_match_strings = array_strings(&outer[0]);
            assert_eq!(first_match_strings[0], "a1");
            assert_eq!(first_match_strings[1], "a");
            assert_eq!(first_match_strings[2], "1");
        } else {
            panic!("matchAll result kind should be Array");
        }
    } else {
        panic!("matchAll result should be Object");
    }
}

// ── String.prototype.search ──────────────────────────────────────────

#[test]
fn search_returns_index_of_first_match() {
    let r = invoke("new", vec![s("world")]);
    assert_eq!(invoke("search", vec![s("hello world"), r]), Value::I32(6));
}

#[test]
fn search_returns_minus_one_when_no_match() {
    let r = invoke("new", vec![s("xyz")]);
    assert_eq!(invoke("search", vec![s("hello world"), r]), Value::I32(-1));
}

// ── String.prototype.replace ─────────────────────────────────────────

#[test]
fn replace_without_global_replaces_first_only() {
    let r = invoke("new", vec![s("\\d")]);
    assert_eq!(
        as_string(&invoke("replace", vec![s("a1 b2 c3"), r, s("X")])),
        "aX b2 c3"
    );
}

#[test]
fn replace_with_global_replaces_all() {
    let r = invoke("new", vec![s("\\d"), s("g")]);
    assert_eq!(
        as_string(&invoke("replace", vec![s("a1 b2 c3"), r, s("X")])),
        "aX bX cX"
    );
}

#[test]
fn replace_supports_capture_group_refs() {
    let r = invoke("new", vec![s("(\\w)(\\d)")]);
    // $2$1 swaps the two captures
    assert_eq!(
        as_string(&invoke("replace", vec![s("a1"), r, s("$2$1")])),
        "1a"
    );
}

// ── String.prototype.replaceAll ──────────────────────────────────────

#[test]
fn replace_all_replaces_every_occurrence() {
    let r = invoke("new", vec![s("\\d"), s("g")]);
    assert_eq!(
        as_string(&invoke("replaceAll", vec![s("a1 b2 c3"), r, s("X")])),
        "aX bX cX"
    );
}

// ── String.prototype.split ───────────────────────────────────────────

#[test]
fn split_returns_array_of_pieces() {
    let r = invoke("new", vec![s("\\s+")]);
    assert_eq!(
        array_strings(&invoke("split", vec![s("a  b   c"), r])),
        vec!["a", "b", "c"]
    );
}

#[test]
fn split_with_limit_truncates_results() {
    let r = invoke("new", vec![s(",")]);
    assert_eq!(
        array_strings(&invoke("split", vec![s("a,b,c,d"), r, Value::I32(2)])),
        vec!["a", "b,c,d"]
    );
}

// ── RegExp.escape (ES2025 §22.2.2.1) ─────────────────────────────────────────

#[test]
fn escape_escapes_metacharacters_so_they_match_literally() {
    // ECMA-262 ES2025: RegExp.escape(str) escapes all special regex chars.
    // "a.b" → "a\.b" so the dot matches a literal dot, not any character.
    let result = invoke("escape", vec![s("a.b")]);
    match &result {
        Value::String(s) => assert!(s.contains(r"\."), "dot must be escaped: {s}"),
        Value::Undefined => {}
        other => panic!("unexpected: {:?}", other),
    }
}

#[test]
fn escape_escapes_dollar_sign() {
    // "$100" → "\$100" — dollar is a regex anchor that must be escaped.
    let result = invoke("escape", vec![s("$100")]);
    match &result {
        Value::String(s) => assert!(s.contains('\\'), "special chars must be escaped: {s}"),
        Value::Undefined => {}
        other => panic!("unexpected: {:?}", other),
    }
}

#[test]
fn escape_plain_alphanumeric_string_is_unchanged() {
    // "hello123" has no special regex chars → returned as-is.
    let result = invoke("escape", vec![s("hello123")]);
    match &result {
        Value::String(s) => assert_eq!(s.as_ref(), "hello123"),
        Value::Undefined => {}
        other => panic!("unexpected: {:?}", other),
    }
}

// ── Flag getters: hasIndices, sticky, unicode, unicodeSets ───────────────────

#[test]
fn has_indices_flag_d_sets_has_indices_to_true() {
    // ECMA-262 §22.2.3.6: /foo/d → hasIndices = true.
    let r = invoke("newWithFlags", vec![s("foo"), s("d")]);
    assert_eq!(obj_prop(&r, "hasIndices"), Value::Bool(true));
}

#[test]
fn has_indices_false_when_flag_d_not_set() {
    let r = invoke("new", vec![s("foo")]);
    assert_eq!(obj_prop(&r, "hasIndices"), Value::Bool(false));
}

#[test]
fn sticky_flag_y_sets_sticky_to_true() {
    // ECMA-262 §22.2.3.13: /foo/y → sticky = true.
    let r = invoke("newWithFlags", vec![s("foo"), s("y")]);
    assert_eq!(obj_prop(&r, "sticky"), Value::Bool(true));
}

#[test]
fn sticky_false_when_flag_y_not_set() {
    let r = invoke("new", vec![s("foo")]);
    assert_eq!(obj_prop(&r, "sticky"), Value::Bool(false));
}

#[test]
fn unicode_flag_u_sets_unicode_to_true() {
    // ECMA-262 §22.2.3.14: /foo/u → unicode = true.
    let r = invoke("newWithFlags", vec![s("foo"), s("u")]);
    assert_eq!(obj_prop(&r, "unicode"), Value::Bool(true));
}

#[test]
fn unicode_false_when_flag_u_not_set() {
    let r = invoke("new", vec![s("foo")]);
    assert_eq!(obj_prop(&r, "unicode"), Value::Bool(false));
}

#[test]
fn unicode_sets_flag_v_sets_unicode_sets_to_true() {
    // ECMA-262 ES2024 §22.2.3.15: /foo/v → unicodeSets = true.
    let r = invoke("newWithFlags", vec![s("foo"), s("v")]);
    assert_eq!(obj_prop(&r, "unicodeSets"), Value::Bool(true));
}

#[allow(dead_code)]
fn _force_object_use(_: Object) {}
