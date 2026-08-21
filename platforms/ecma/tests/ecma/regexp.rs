//! Behaviour tests for `ecma:regexp` host imports — ECMA-262 §22.2
//! RegExp constructor + `RegExp.prototype.{test, exec, toString}` plus
//! the regex-taking `String.prototype.{match, matchAll, search, replace,
//! replaceAll, split}` methods.
//!
//! Reference: ECMA-262 §22.2 RegExp + §22.1.3.{13,14,16,18,19,20}.

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    invoke_result(name, args).expect("VM run failed")
}

fn invoke_result(name: &str, args: Vec<Value>) -> Result<Value, vybe_runtime::VMError> {
    let (result, _) = invoke_result_with_exception(name, args);
    result
}

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn push_arg(vm: &mut VM, chunk: &mut Chunk, value: Value) {
    match value {
        Value::I32(n) => chunk.emit_i32_const(n, 0),
        Value::I64(n) => chunk.emit_i64_const(n, 0),
        Value::F32(f) => chunk.emit_f32_const(f, 0),
        Value::F64(f) => chunk.emit_f64_const(f, 0),
        Value::Bool(b) => chunk.emit_bool_const(b, 0),
        Value::String(s) => chunk.emit_string_const(&s, 0),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
        other => {
            let global = format!(
                "__test_arg_{}",
                TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
            );
            vm.set_global_owned(global.clone(), other);
            let ci = chunk.intern_string_constant(&global);
            chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
        }
    }
}

fn invoke_result_with_exception(
    name: &str,
    args: Vec<Value>,
) -> (Result<Value, vybe_runtime::VMError>, Option<Value>) {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-regexp-test>");
    let import_idx = chunk.add_import("ecma:regexp", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let result = vm.run(vec![chunk]);
    let exception = vm.last_exception.clone();
    (result, exception)
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
        return object
            .properties
            .get(key)
            .cloned()
            .unwrap_or(Value::Undefined);
    }
    Value::Undefined
}

fn assert_throws_error_name(name: &str, args: Vec<Value>, expected_name: &str) {
    let (result, exception) = invoke_result_with_exception(name, args);
    result.expect_err("host call should throw");
    let exception = exception.expect("host call should preserve thrown value");
    assert_eq!(as_string(&obj_prop(&exception, "name")), expected_name);
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

#[test]
fn new_rejects_u_and_v_flags_together() {
    assert_throws_error_name("new", vec![s("foo"), s("uv")], "SyntaxError");
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
    assert_eq!(
        invoke("test", vec![r, s("hello world")]),
        Value::Bool(false)
    );
}

#[test]
fn test_with_case_insensitive_flag() {
    let r = invoke("new", vec![s("HELLO"), s("i")]);
    assert_eq!(invoke("test", vec![r, s("hello world")]), Value::Bool(true));
}

#[test]
fn new_rejects_invalid_pattern() {
    assert_throws_error_name("new", vec![s("[")], "SyntaxError");
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
    let r = invoke("new", vec![s("(\\w)(\\d)"), s("g")]);
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

#[test]
fn match_all_with_non_global_regexp_throws_type_error() {
    let r = invoke("new", vec![s("(\\w)(\\d)")]);
    assert_throws_error_name("matchAll", vec![s("a1 b2"), r], "TypeError");
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

#[test]
fn replace_supports_named_capture_group_refs() {
    let r = invoke(
        "new",
        vec![s("(?<year>\\d{4})-(?<month>\\d{2})-(?<day>\\d{2})")],
    );
    assert_eq!(
        as_string(&invoke(
            "replace",
            vec![s("2024-06-15"), r, s("$<day>/$<month>/$<year>")]
        )),
        "15/06/2024"
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
        vec!["a", "b"]
    );
}

#[test]
fn split_with_regex_limit_counts_captures_and_truncates() {
    let r = invoke("new", vec![s("([|,])")]);
    assert_eq!(
        array_strings(&invoke("split", vec![s("a|b,c"), r, Value::I32(3)])),
        vec!["a", "|", "b"]
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

// ── Unicode property escapes — ECMA-262 §22.2.2.9 ────────────────────────────
//
// Tests cover all three spec tables:
//   • table-binary-unicode-properties.html       (u flag, \p{Name})
//   • table-nonbinary-unicode-properties.html    (u flag, \p{Name=Value})
//   • table-binary-unicode-properties-of-strings.html (v flag, \p{Name})

fn matches_u(pattern: &str, input: &str) -> bool {
    let re = invoke("newWithFlags", vec![s(pattern), s("u")]);
    invoke("test", vec![re, s(input)]).as_i32() != 0
}

fn matches_v(pattern: &str, input: &str) -> bool {
    let re = invoke("newWithFlags", vec![s(pattern), s("v")]);
    invoke("test", vec![re, s(input)]).as_i32() != 0
}

// ── Binary properties — ASCII / ASCII_Hex_Digit ───────────────────────────────

#[test]
fn binary_ascii_matches_basic_latin() {
    assert!(matches_u(r"\p{ASCII}", "A"));
    assert!(matches_u(r"\p{ASCII}", "z"));
    assert!(matches_u(r"\p{ASCII}", "0"));
}

#[test]
fn binary_ascii_rejects_non_ascii() {
    assert!(!matches_u(r"\p{ASCII}", "é"));
    assert!(!matches_u(r"\p{ASCII}", "😀"));
    assert!(!matches_u(r"\p{ASCII}", "日"));
}

#[test]
fn binary_ascii_hex_digit_matches_hex_chars() {
    // §22.2.2.9: AHex alias
    assert!(matches_u(r"\p{ASCII_Hex_Digit}", "A"));
    assert!(matches_u(r"\p{ASCII_Hex_Digit}", "f"));
    assert!(matches_u(r"\p{ASCII_Hex_Digit}", "9"));
    assert!(matches_u(r"\p{AHex}", "B")); // canonical alias
}

#[test]
fn binary_ascii_hex_digit_rejects_non_hex() {
    assert!(!matches_u(r"\p{ASCII_Hex_Digit}", "G"));
    assert!(!matches_u(r"\p{ASCII_Hex_Digit}", "z"));
}

// ── Binary properties — Alpha / Alphabetic ───────────────────────────────────

#[test]
fn binary_alpha_matches_letters() {
    assert!(matches_u(r"\p{Alpha}", "a"));
    assert!(matches_u(r"\p{Alpha}", "Z"));
    assert!(matches_u(r"\p{Alpha}", "é")); // Latin extended
    assert!(matches_u(r"\p{Alpha}", "α")); // Greek
    assert!(matches_u(r"\p{Alphabetic}", "a")); // canonical alias
}

#[test]
fn binary_alpha_rejects_digits_and_punctuation() {
    assert!(!matches_u(r"\p{Alpha}", "1"));
    assert!(!matches_u(r"\p{Alpha}", "!"));
    assert!(!matches_u(r"\p{Alpha}", " "));
}

// ── Binary properties — Uppercase / Lowercase ────────────────────────────────

#[test]
fn binary_uppercase_matches_uppercase_letters() {
    assert!(matches_u(r"\p{Uppercase}", "A"));
    assert!(matches_u(r"\p{Uppercase}", "É"));
    assert!(matches_u(r"\p{Upper}", "Z")); // alias
}

#[test]
fn binary_lowercase_matches_lowercase_letters() {
    assert!(matches_u(r"\p{Lowercase}", "a"));
    assert!(matches_u(r"\p{Lowercase}", "é"));
    assert!(matches_u(r"\p{Lower}", "z")); // alias
}

#[test]
fn binary_uppercase_and_lowercase_are_disjoint_for_ascii() {
    assert!(!matches_u(r"\p{Uppercase}", "a"));
    assert!(!matches_u(r"\p{Lowercase}", "A"));
}

// ── Binary properties — White_Space ──────────────────────────────────────────

#[test]
fn binary_white_space_matches_whitespace_chars() {
    assert!(matches_u(r"\p{White_Space}", " "));
    assert!(matches_u(r"\p{White_Space}", "\t"));
    assert!(matches_u(r"\p{White_Space}", "\n"));
    assert!(matches_u(r"\p{space}", "  ")); // alias
}

#[test]
fn binary_white_space_rejects_non_whitespace() {
    assert!(!matches_u(r"\p{White_Space}", "a"));
    assert!(!matches_u(r"\p{White_Space}", "1"));
}

// ── Binary properties — Emoji / Extended_Pictographic ────────────────────────

#[test]
fn binary_emoji_matches_emoji_chars() {
    assert!(matches_u(r"\p{Emoji}", "😀"));
    assert!(matches_u(r"\p{Emoji}", "❤"));
}

#[test]
fn binary_emoji_rejects_plain_ascii() {
    assert!(!matches_u(r"\p{Emoji}", "a"));
    assert!(!matches_u(r"\p{Emoji}", "!")); // punctuation has Emoji=No (digits are Emoji=Yes for keycaps)
}

#[test]
fn binary_extended_pictographic_matches_pictographs() {
    assert!(matches_u(r"\p{Extended_Pictographic}", "😀"));
    assert!(matches_u(r"\p{ExtPict}", "🎉")); // alias
}

// ── Binary properties — ID_Start / ID_Continue ───────────────────────────────

#[test]
fn binary_id_start_matches_identifier_start_chars() {
    // §22.2.2.9: ID_Start covers letters and underscore-like chars
    assert!(matches_u(r"\p{ID_Start}", "a"));
    assert!(matches_u(r"\p{ID_Start}", "Z"));
    assert!(matches_u(r"\p{ID_Start}", "α"));
}

#[test]
fn binary_id_start_rejects_digits_and_space() {
    assert!(!matches_u(r"\p{ID_Start}", "1"));
    assert!(!matches_u(r"\p{ID_Start}", " "));
}

#[test]
fn binary_id_continue_includes_digits_after_start() {
    assert!(matches_u(r"\p{ID_Continue}", "a"));
    assert!(matches_u(r"\p{ID_Continue}", "1")); // digits are valid mid-identifier
    assert!(matches_u(r"\p{XIDC}", "a")); // alias
    assert!(matches_u(r"\p{XIDS}", "a")); // XID_Start alias
}

// ── Binary properties — Math / Dash / Diacritic ──────────────────────────────

#[test]
fn binary_math_matches_math_symbols() {
    assert!(matches_u(r"\p{Math}", "+"));
    assert!(matches_u(r"\p{Math}", "="));
    assert!(matches_u(r"\p{Math}", "×"));
}

#[test]
fn binary_dash_matches_hyphen_and_dashes() {
    assert!(matches_u(r"\p{Dash}", "-"));
    assert!(matches_u(r"\p{Dash}", "—")); // em dash
}

#[test]
fn binary_diacritic_matches_combining_marks() {
    // Combining acute accent (U+0301) has Diacritic=Yes
    assert!(matches_u(r"\p{Diacritic}", "\u{0301}"));
    assert!(matches_u(r"\p{Dia}", "\u{0300}")); // alias
}

// ── Binary properties — Pattern_White_Space / Pattern_Syntax ─────────────────

#[test]
fn binary_pattern_white_space_matches_spec_whitespace() {
    assert!(matches_u(r"\p{Pattern_White_Space}", " "));
    assert!(matches_u(r"\p{Pat_WS}", "\t")); // alias
}

#[test]
fn binary_pattern_syntax_matches_syntax_chars() {
    assert!(matches_u(r"\p{Pattern_Syntax}", "!"));
    assert!(matches_u(r"\p{Pat_Syn}", "{")); // alias
}

// ── Binary properties — Bidi_Control ─────────────────────────────────────────

#[test]
fn binary_bidi_control_matches_bidi_format_chars() {
    // U+200F RIGHT-TO-LEFT MARK has Bidi_Control=Yes
    assert!(matches_u(r"\p{Bidi_Control}", "\u{200F}"));
    assert!(matches_u(r"\p{Bidi_C}", "\u{200E}")); // alias — LRM
}

// ── Non-binary properties: General_Category ──────────────────────────────────

#[test]
fn general_category_lu_matches_uppercase_letters() {
    assert!(matches_u(r"\p{General_Category=Lu}", "A"));
    assert!(matches_u(r"\p{gc=Lu}", "É")); // alias
}

#[test]
fn general_category_ll_matches_lowercase_letters() {
    assert!(matches_u(r"\p{General_Category=Ll}", "a"));
    assert!(matches_u(r"\p{gc=Ll}", "é"));
}

#[test]
fn general_category_nd_matches_decimal_digits() {
    // Nd = Decimal Number
    assert!(matches_u(r"\p{General_Category=Nd}", "0"));
    assert!(matches_u(r"\p{gc=Nd}", "9"));
    assert!(!matches_u(r"\p{gc=Nd}", "a"));
}

#[test]
fn general_category_z_matches_separator_category() {
    // Zs = Space_Separator
    assert!(matches_u(r"\p{gc=Zs}", " "));
    // P = Punctuation (supercat)
    assert!(matches_u(r"\p{gc=P}", "."));
}

// ── Non-binary properties: Script ────────────────────────────────────────────

#[test]
fn script_latin_matches_latin_letters() {
    assert!(matches_u(r"\p{Script=Latin}", "a"));
    assert!(matches_u(r"\p{Script=Latin}", "é"));
    assert!(matches_u(r"\p{sc=Latin}", "Z")); // alias
}

#[test]
fn script_latin_rejects_non_latin() {
    assert!(!matches_u(r"\p{Script=Latin}", "α")); // Greek
    assert!(!matches_u(r"\p{Script=Latin}", "あ")); // Hiragana
}

#[test]
fn script_greek_matches_greek_letters() {
    assert!(matches_u(r"\p{Script=Greek}", "α"));
    assert!(matches_u(r"\p{Script=Greek}", "Ω"));
    assert!(!matches_u(r"\p{Script=Greek}", "a"));
}

#[test]
fn script_hiragana_matches_hiragana() {
    assert!(matches_u(r"\p{Script=Hiragana}", "あ"));
    assert!(matches_u(r"\p{sc=Hiragana}", "の"));
    assert!(!matches_u(r"\p{Script=Hiragana}", "a"));
}

#[test]
fn script_cyrillic_matches_cyrillic() {
    assert!(matches_u(r"\p{Script=Cyrillic}", "а")); // Cyrillic а
    assert!(!matches_u(r"\p{Script=Cyrillic}", "a")); // Latin a
}

#[test]
fn script_han_matches_cjk_ideographs() {
    assert!(matches_u(r"\p{Script=Han}", "日"));
    assert!(matches_u(r"\p{sc=Han}", "中"));
}

// ── Non-binary properties: Script_Extensions ─────────────────────────────────

#[test]
fn script_extensions_matches_chars_used_in_multiple_scripts() {
    // U+0300 COMBINING GRAVE ACCENT is used in many scripts (Latin, Greek…)
    assert!(matches_u(r"\p{Script_Extensions=Latin}", "\u{0300}"));
    assert!(matches_u(r"\p{scx=Greek}", "\u{0300}"));
}

// ── String binary properties (v flag) ────────────────────────────────────────

#[test]
fn string_property_rgi_emoji_matches_emoji() {
    assert!(matches_v(r"\p{RGI_Emoji}", "😀"));
    assert!(matches_v(r"\p{RGI_Emoji}", "❤️"));
}

#[test]
fn string_property_rgi_emoji_rejects_ascii() {
    assert!(!matches_v(r"\p{RGI_Emoji}", "a"));
}

#[test]
fn string_property_basic_emoji_matches_single_emoji() {
    assert!(matches_v(r"\p{Basic_Emoji}", "😀"));
}

#[test]
fn string_property_emoji_keycap_sequence_matches_keycaps() {
    // Keycap sequences: digit + U+FE0F + U+20E3
    assert!(matches_v(r"\p{Emoji_Keycap_Sequence}", "1️⃣"));
}

// ── Property negation \P{} ────────────────────────────────────────────────────

#[test]
fn negated_property_ascii_rejects_ascii_chars() {
    assert!(!matches_u(r"\P{ASCII}", "A"));
    assert!(matches_u(r"\P{ASCII}", "é"));
    assert!(matches_u(r"\P{ASCII}", "😀"));
}

#[test]
fn negated_property_uppercase_matches_lowercase() {
    assert!(matches_u(r"\P{Uppercase}", "a"));
    assert!(!matches_u(r"\P{Uppercase}", "A"));
}

// ── Properties in character classes ─────────────────────────────────────────

#[test]
fn property_in_character_class_with_other_chars() {
    // [a-z\p{Uppercase}] matches lowercase or uppercase
    assert!(matches_u(r"[a-z\p{Uppercase}]", "A"));
    assert!(matches_u(r"[a-z\p{Uppercase}]", "z"));
    assert!(!matches_u(r"[a-z\p{Uppercase}]", "1"));
}

#[test]
fn negated_property_in_character_class() {
    // [\P{ASCII}] matches non-ASCII only
    assert!(matches_u(r"[\P{ASCII}]", "é"));
    assert!(!matches_u(r"[\P{ASCII}]", "a"));
}

// ── v-flag set operations with properties ────────────────────────────────────

#[test]
fn v_flag_union_of_properties_with_set_syntax() {
    // [\p{Alpha}&&\p{ASCII}] = ASCII letters (intersection)
    assert!(matches_v(r"[\p{Alpha}&&\p{ASCII}]", "a"));
    assert!(!matches_v(r"[\p{Alpha}&&\p{ASCII}]", "é")); // Alpha but not ASCII
    assert!(!matches_v(r"[\p{Alpha}&&\p{ASCII}]", "1")); // ASCII but not Alpha
}

#[test]
fn v_flag_difference_of_properties() {
    // [\p{ASCII}--\p{Alpha}] = ASCII non-letters (digits, punct, etc.)
    assert!(matches_v(r"[\p{ASCII}--\p{Alpha}]", "1"));
    assert!(matches_v(r"[\p{ASCII}--\p{Alpha}]", "!"));
    assert!(!matches_v(r"[\p{ASCII}--\p{Alpha}]", "a")); // is Alpha
}
