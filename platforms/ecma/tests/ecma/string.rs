//! Behaviour tests for `ecma:string` host imports — the ECMA-262
//! `String.prototype` and `String` static surface, exposed as the
//! canonical JS-runtime string ops Vybe-emitted .wasm calls into.
//!
//! Reference: ECMA-262 §22.1 Strings + String.prototype.
//!
//! Where the merged `wasm:js-string` proposal already covers an
//! op (`length`, `concat`, `charCodeAt`, `substring`, `equals`,
//! `compare`, `fromCharCode`, `fromCodePoint`), `ecma:string`
//! delegates rather than reimplementing — keeping a single source
//! of truth and maximising cross-runtime portability.

use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

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

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-string-test>");
    let import_idx = chunk.add_import("ecma:string", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.run(vec![chunk]).expect("VM run failed")
}

fn invoke_value(name: &str, args: Vec<Value>) -> Value {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let mut chunk = Chunk::new("<ecma-value-test>");
    let import_idx = chunk.add_import("ecma:value", name);
    let argc = args.len() as u8;
    for value in args {
        push_arg(&mut vm, &mut chunk, value);
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

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

// ── length / charAt / charCodeAt / codePointAt ────────────────────

#[test]
fn length_counts_utf16_code_units() {
    assert_eq!(invoke("length", vec![s("hello")]), Value::F64(5.0));
    assert_eq!(invoke("length", vec![s("")]), Value::F64(0.0));
}

#[test]
fn char_at_returns_single_character() {
    assert_eq!(
        as_string(&invoke("charAt", vec![s("hello"), Value::F64(1.0)])),
        "e"
    );
}

#[test]
fn char_at_out_of_range_returns_empty_string() {
    // ECMA-262: out-of-range returns empty string (NOT undefined).
    assert_eq!(
        as_string(&invoke("charAt", vec![s("hi"), Value::F64(99.0)])),
        ""
    );
}

#[test]
fn char_code_at_returns_unit_code() {
    assert_eq!(
        invoke("charCodeAt", vec![s("Aa"), Value::F64(0.0)]),
        Value::F64(65.0)
    );
    assert_eq!(
        invoke("charCodeAt", vec![s("Aa"), Value::F64(1.0)]),
        Value::F64(97.0)
    );
}

#[test]
fn at_supports_negative_index() {
    // ECMA-262 §22.1.3.1 (added in ES2022).
    assert_eq!(
        as_string(&invoke("at", vec![s("hello"), Value::F64(-1.0)])),
        "o"
    );
    assert_eq!(
        as_string(&invoke("at", vec![s("hello"), Value::F64(0.0)])),
        "h"
    );
}

// ── concat ────────────────────────────────────────────────────────

#[test]
fn concat_joins_arguments() {
    assert_eq!(
        as_string(&invoke("concat", vec![s("ab"), s("cd"), s("ef")])),
        "abcdef"
    );
}

// ── substring / slice ─────────────────────────────────────────────

#[test]
fn substring_extracts_range() {
    // substring(start, end) — both clamped non-negative.
    assert_eq!(
        as_string(&invoke(
            "substring",
            vec![s("abcdef"), Value::F64(1.0), Value::F64(4.0)]
        )),
        "bcd"
    );
}

#[test]
fn substring_swaps_when_start_greater_than_end() {
    // ECMA-262: substring swaps args if start > end.
    assert_eq!(
        as_string(&invoke(
            "substring",
            vec![s("abcdef"), Value::F64(4.0), Value::F64(1.0)]
        )),
        "bcd"
    );
}

#[test]
fn slice_extracts_range() {
    assert_eq!(
        as_string(&invoke(
            "slice",
            vec![s("abcdef"), Value::F64(1.0), Value::F64(4.0)]
        )),
        "bcd"
    );
}

#[test]
fn slice_supports_negative_indices() {
    // slice — negative indices count from end (substring does not).
    assert_eq!(
        as_string(&invoke("slice", vec![s("abcdef"), Value::F64(-3.0)])),
        "def"
    );
    assert_eq!(
        as_string(&invoke(
            "slice",
            vec![s("abcdef"), Value::F64(-3.0), Value::F64(-1.0)]
        )),
        "de"
    );
}

// ── includes / indexOf / lastIndexOf ──────────────────────────────

#[test]
fn includes_finds_substring() {
    assert_eq!(
        invoke("includes", vec![s("hello world"), s("world")]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("includes", vec![s("hello world"), s("XYZ")]),
        Value::Bool(false)
    );
}

#[test]
fn includes_with_start_position() {
    assert_eq!(
        invoke(
            "includes",
            vec![s("hello world"), s("hello"), Value::F64(1.0)]
        ),
        Value::Bool(false)
    );
}

#[test]
fn index_of_returns_position_or_minus_one() {
    assert_eq!(
        invoke("indexOf", vec![s("hello world"), s("world")]),
        Value::F64(6.0)
    );
    assert_eq!(
        invoke("indexOf", vec![s("hello world"), s("XYZ")]),
        Value::F64(-1.0)
    );
}

#[test]
fn last_index_of_finds_last_occurrence() {
    assert_eq!(
        invoke("lastIndexOf", vec![s("ababab"), s("ab")]),
        Value::F64(4.0)
    );
    assert_eq!(
        invoke("lastIndexOf", vec![s("ababab"), s("XYZ")]),
        Value::F64(-1.0)
    );
}

#[test]
fn last_index_of_honors_start_position() {
    assert_eq!(
        invoke("lastIndexOf", vec![s("ababa"), s("ba"), Value::F64(2.0)]),
        Value::F64(1.0)
    );
    assert_eq!(
        invoke("lastIndexOf", vec![s("banana"), s("a"), Value::F64(4.0)]),
        Value::F64(3.0)
    );
    assert_eq!(
        invoke("lastIndexOf", vec![s("abcabc"), s("a"), Value::F64(-2.0)]),
        Value::F64(0.0)
    );
    assert_eq!(
        invoke("lastIndexOf", vec![s("abc"), s(""), Value::F64(1.0)]),
        Value::F64(1.0)
    );
}

#[test]
fn last_index_of_position_uses_to_integer_or_infinity() {
    assert_eq!(
        invoke("lastIndexOf", vec![s("ababa"), s("ba"), Value::Undefined]),
        Value::F64(3.0)
    );
    assert_eq!(
        invoke(
            "lastIndexOf",
            vec![s("ababa"), s("ba"), Value::F64(f64::NAN)]
        ),
        Value::F64(-1.0)
    );
    assert_eq!(
        invoke(
            "lastIndexOf",
            vec![s("ababa"), s("ba"), Value::F64(f64::INFINITY)]
        ),
        Value::F64(3.0)
    );
    assert_eq!(
        invoke("lastIndexOf", vec![s("ababa"), s("ba"), s("2")]),
        Value::F64(1.0)
    );
    assert_eq!(
        invoke("lastIndexOf", vec![s("ababa"), s("ba"), Value::Bool(true)]),
        Value::F64(1.0)
    );
}

#[test]
fn last_index_of_uses_utf16_code_unit_indices() {
    assert_eq!(
        invoke("lastIndexOf", vec![s("😀a😀"), s("😀")]),
        Value::F64(3.0)
    );
    assert_eq!(
        invoke("lastIndexOf", vec![s("😀a😀"), s("😀"), Value::F64(2.0)]),
        Value::F64(0.0)
    );
}

// ── startsWith / endsWith ─────────────────────────────────────────

#[test]
fn starts_with_returns_bool() {
    assert_eq!(
        invoke("startsWith", vec![s("hello"), s("hel")]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("startsWith", vec![s("hello"), s("ell")]),
        Value::Bool(false)
    );
}

#[test]
fn ends_with_returns_bool() {
    assert_eq!(
        invoke("endsWith", vec![s("hello"), s("llo")]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("endsWith", vec![s("hello"), s("hel")]),
        Value::Bool(false)
    );
}

// ── toUpperCase / toLowerCase ─────────────────────────────────────

#[test]
fn to_upper_case_capitalises() {
    assert_eq!(as_string(&invoke("toUpperCase", vec![s("hello")])), "HELLO");
}

#[test]
fn to_lower_case_uncapitalises() {
    assert_eq!(as_string(&invoke("toLowerCase", vec![s("HELLO")])), "hello");
}

// ── trim / trimStart / trimEnd ────────────────────────────────────

#[test]
fn trim_strips_whitespace_both_sides() {
    assert_eq!(as_string(&invoke("trim", vec![s("  hello  ")])), "hello");
}

#[test]
fn trim_start_strips_only_leading() {
    assert_eq!(
        as_string(&invoke("trimStart", vec![s("  hello  ")])),
        "hello  "
    );
}

#[test]
fn trim_end_strips_only_trailing() {
    assert_eq!(
        as_string(&invoke("trimEnd", vec![s("  hello  ")])),
        "  hello"
    );
}

// ── padStart / padEnd ─────────────────────────────────────────────

#[test]
fn pad_start_pads_to_target_length() {
    assert_eq!(
        as_string(&invoke("padStart", vec![s("5"), Value::F64(3.0), s("0")])),
        "005"
    );
}

#[test]
fn pad_start_default_pad_is_space() {
    assert_eq!(
        as_string(&invoke("padStart", vec![s("hi"), Value::F64(5.0)])),
        "   hi"
    );
}

#[test]
fn pad_end_pads_to_target_length() {
    assert_eq!(
        as_string(&invoke("padEnd", vec![s("5"), Value::F64(3.0), s(".")])),
        "5.."
    );
}

// ── repeat ────────────────────────────────────────────────────────

#[test]
fn repeat_concatenates_n_times() {
    assert_eq!(
        as_string(&invoke("repeat", vec![s("ab"), Value::F64(3.0)])),
        "ababab"
    );
}

#[test]
fn repeat_zero_returns_empty() {
    assert_eq!(
        as_string(&invoke("repeat", vec![s("ab"), Value::F64(0.0)])),
        ""
    );
}

// ── replace / replaceAll ──────────────────────────────────────────

#[test]
fn replace_replaces_first_occurrence_only() {
    // ECMA-262 String.prototype.replace with a string searchValue
    // replaces only the FIRST occurrence (regex w/ /g flag is
    // different — that's matchAll territory).
    assert_eq!(
        as_string(&invoke("replace", vec![s("ababab"), s("ab"), s("X")])),
        "Xabab"
    );
}

#[test]
fn replace_all_replaces_every_occurrence() {
    assert_eq!(
        as_string(&invoke("replaceAll", vec![s("ababab"), s("ab"), s("X")])),
        "XXX"
    );
}

// ── split ─────────────────────────────────────────────────────────

#[test]
fn split_returns_array_of_pieces() {
    assert_eq!(
        array_strings(&invoke("split", vec![s("a,b,c"), s(",")])),
        vec!["a", "b", "c"]
    );
}

#[test]
fn split_with_limit_truncates() {
    assert_eq!(
        array_strings(&invoke(
            "split",
            vec![s("a,b,c,d"), s(","), Value::F64(2.0)]
        )),
        vec!["a", "b"]
    );
}

#[test]
fn split_empty_separator_yields_per_character() {
    assert_eq!(
        array_strings(&invoke("split", vec![s("abc"), s("")])),
        vec!["a", "b", "c"]
    );
}

// ── String.fromCharCode / fromCodePoint (constructor statics) ─────

#[test]
fn from_char_code_builds_string() {
    assert_eq!(
        as_string(&invoke(
            "fromCharCode",
            vec![Value::F64(65.0), Value::F64(97.0)]
        )),
        "Aa"
    );
}

#[test]
fn from_code_point_builds_string() {
    // 0x1F600 is 😀 — outside BMP, encoded as surrogate pair in UTF-16.
    let v = invoke("fromCodePoint", vec![Value::F64(0x1F600 as f64)]);
    assert_eq!(as_string(&v), "\u{1F600}");
}

// ── valueOf ───────────────────────────────────────────────────────

#[test]
fn value_of_returns_underlying_string() {
    assert_eq!(as_string(&invoke("valueOf", vec![s("hi")])), "hi");
}

// ── localeCompare (ECMA-262 §22.1.3.10) ─────────────────────────────

#[test]
fn locale_compare_returns_negative_for_less() {
    if let Value::I32(n) = invoke("localeCompare", vec![s("a"), s("b")]) {
        assert!(n < 0);
    } else {
        panic!("localeCompare should return I32");
    }
}

#[test]
fn locale_compare_returns_zero_for_equal() {
    assert_eq!(
        invoke("localeCompare", vec![s("abc"), s("abc")]),
        Value::I32(0)
    );
}

#[test]
fn locale_compare_returns_positive_for_greater() {
    if let Value::I32(n) = invoke("localeCompare", vec![s("z"), s("a")]) {
        assert!(n > 0);
    } else {
        panic!("localeCompare should return I32");
    }
}

#[test]
fn value_invoke_method_locale_compare_on_string_primitive() {
    assert_eq!(
        invoke_value("invokeMethod", vec![s("a"), s("localeCompare"), s("b")]),
        Value::I32(-1)
    );
}

// ── normalize (ECMA-262 §22.1.3.13) ─────────────────────────────────

#[test]
fn normalize_passes_ascii_through_unchanged() {
    // ASCII is already in NFC form so normalize is identity.
    assert_eq!(as_string(&invoke("normalize", vec![s("hello")])), "hello");
}

// ── ECMA-262 §19.2.6 URI globals ────────────────────────────────────

#[test]
fn encode_uri_component_escapes_reserved_chars() {
    // Spaces, slashes, &, ?, # all get %-encoded by encodeURIComponent
    assert_eq!(
        as_string(&invoke("encodeURIComponent", vec![s("hello world")])),
        "hello%20world"
    );
    assert_eq!(
        as_string(&invoke("encodeURIComponent", vec![s("a&b=c")])),
        "a%26b%3Dc"
    );
}

#[test]
fn encode_uri_component_leaves_unreserved_alone() {
    // ECMA-262 §19.2.6.5: unreserved set is alpha + digit + -_.!~*'()
    assert_eq!(
        as_string(&invoke("encodeURIComponent", vec![s("abc123-_.!~*'()")])),
        "abc123-_.!~*'()"
    );
}

#[test]
fn decode_uri_component_reverses_encode() {
    assert_eq!(
        as_string(&invoke("decodeURIComponent", vec![s("hello%20world")])),
        "hello world"
    );
}

#[test]
fn encode_uri_preserves_uri_syntax_chars() {
    // ECMA-262 §19.2.6.4: encodeURI leaves `;,/?:@&=+$#` unencoded
    // unlike encodeURIComponent.
    assert_eq!(
        as_string(&invoke(
            "encodeURI",
            vec![s("https://example.com/path?q=1")]
        )),
        "https://example.com/path?q=1"
    );
}

#[test]
fn encode_uri_still_escapes_spaces() {
    // Space isn't in encodeURI's preserved set; gets percent-encoded.
    assert_eq!(
        as_string(&invoke("encodeURI", vec![s("hello world")])),
        "hello%20world"
    );
}

#[test]
fn decode_uri_reverses_encode_uri() {
    assert_eq!(
        as_string(&invoke("decodeURI", vec![s("hello%20world")])),
        "hello world"
    );
}

// ── WHATWG btoa / atob (HTML §8.3) ──────────────────────────────────

#[test]
fn btoa_encodes_to_base64() {
    assert_eq!(as_string(&invoke("btoa", vec![s("hello")])), "aGVsbG8=");
}

#[test]
fn atob_decodes_from_base64() {
    assert_eq!(as_string(&invoke("atob", vec![s("aGVsbG8=")])), "hello");
}

#[test]
fn btoa_atob_roundtrip() {
    let encoded = invoke("btoa", vec![s("The quick brown fox")]);
    let decoded = invoke("atob", vec![encoded]);
    assert_eq!(as_string(&decoded), "The quick brown fox");
}

// ── codePointAt (ECMA-262 §22.1.3.2) ───────────────────────────────────────

#[test]
fn code_point_at_returns_full_unicode_codepoint() {
    // "A" = U+0041 = 65. This agrees with charCodeAt for BMP characters,
    // but codePointAt is the spec-correct API for code-point access.
    assert_eq!(
        invoke("codePointAt", vec![s("A"), Value::F64(0.0)]),
        Value::F64(65.0)
    );
}

#[test]
fn code_point_at_out_of_bounds_returns_undefined() {
    assert_eq!(
        invoke("codePointAt", vec![s("hi"), Value::F64(5.0)]),
        Value::Undefined
    );
}

// ── match (ECMA-262 §22.1.3.12) ─────────────────────────────────────────────

#[test]
fn match_with_pattern_returns_first_match_array() {
    // Without /g flag, match returns the first match + capture groups.
    let result = invoke("match", vec![s("hello world"), s("\\w+")]);
    // Must be an array whose first element is the matched string.
    if let Value::Object(o) = result {
        if let ObjectKind::Array(elems) = &o.lock().unwrap().kind {
            assert_eq!(elems.first().cloned(), Some(s("hello")));
        } else {
            panic!("expected array kind");
        }
    } else {
        panic!("expected array");
    }
}

#[test]
fn match_with_no_hit_returns_null() {
    // ECMA-262: match returns null when there is no match.
    assert_eq!(invoke("match", vec![s("hello"), s("\\d+")]), Value::Null);
}

// ── search (ECMA-262 §22.1.3.21) ────────────────────────────────────────────

#[test]
fn search_returns_index_of_first_match() {
    assert_eq!(
        invoke("search", vec![s("hello world"), s("world")]),
        Value::F64(6.0)
    );
}

#[test]
fn search_returns_negative_one_for_no_match() {
    assert_eq!(
        invoke("search", vec![s("hello"), s("\\d+")]),
        Value::F64(-1.0)
    );
}

// ── startsWith / endsWith with position parameter ───────────────────────────

#[test]
fn starts_with_with_position_skips_prefix_chars() {
    // startsWith("hello", "ello", 1) — searching from index 1.
    assert_eq!(
        invoke("startsWith", vec![s("hello"), s("ello"), Value::F64(1.0)]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("startsWith", vec![s("hello"), s("hello"), Value::F64(1.0)]),
        Value::Bool(false)
    );
}

#[test]
fn ends_with_with_end_position_treats_string_as_ending_earlier() {
    // endsWith("hello", "hell", 4) — acts as if the string has length 4 → "hell".
    assert_eq!(
        invoke("endsWith", vec![s("hello"), s("hell"), Value::F64(4.0)]),
        Value::Bool(true)
    );
    assert_eq!(
        invoke("endsWith", vec![s("hello"), s("ello"), Value::F64(4.0)]),
        Value::Bool(false)
    );
}

// ── padStart — no-op when string already at target length ───────────────────

#[test]
fn pad_start_is_noop_when_string_longer_than_target() {
    // padStart(1, "0") on a 3-char string must not truncate.
    assert_eq!(
        as_string(&invoke("padStart", vec![s("abc"), Value::F64(1.0), s("0")])),
        "abc"
    );
}

// ── String.raw (ECMA-262 §22.1.2.4) ─────────────────────────────────────────

#[test]
fn raw_interpolates_without_escape_processing() {
    // String.raw receives a template-object (array of raw strings) + substitutions.
    // The host function takes the raw strings array + subs and joins them.
    let raw_parts = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(
        vybe_runtime::value::Object::new_array(vec![s("Hello\\n"), s("!")]),
    )));
    let result = invoke("raw", vec![raw_parts, s("World")]);
    // Result should be "Hello\nWorld!" with a literal backslash-n, not a newline.
    assert_eq!(as_string(&result), "Hello\\nWorld!");
}

// ── isWellFormed / toWellFormed (ES2024 §22.1.3.14 / §22.1.3.33) ─────────────

#[test]
fn is_well_formed_returns_true_for_valid_utf16_string() {
    // ECMA-262 ES2024: "hello" has no lone surrogates → isWellFormed returns true.
    assert_eq!(invoke("isWellFormed", vec![s("hello")]), Value::Bool(true));
}

#[test]
fn is_well_formed_returns_true_for_empty_string() {
    assert_eq!(invoke("isWellFormed", vec![s("")]), Value::Bool(true));
}

#[test]
fn to_well_formed_returns_well_formed_string_unchanged() {
    // ECMA-262 ES2024: toWellFormed on a well-formed string returns it unchanged.
    let result = as_string(&invoke("toWellFormed", vec![s("hello")]));
    assert_eq!(result, "hello");
}

// ── toLocaleUpperCase / toLocaleLowerCase ─────────────────────────────────────

#[test]
fn to_locale_upper_case_uppercases_ascii() {
    // ECMA-262 §22.1.3.27: toLocaleUpperCase is locale-sensitive; for ASCII it matches toUpperCase.
    assert_eq!(
        as_string(&invoke("toLocaleUpperCase", vec![s("hello")])),
        "HELLO"
    );
}

#[test]
fn to_locale_lower_case_lowercases_ascii() {
    // ECMA-262 §22.1.3.26: toLocaleLowerCase is locale-sensitive; for ASCII it matches toLowerCase.
    assert_eq!(
        as_string(&invoke("toLocaleLowerCase", vec![s("WORLD")])),
        "world"
    );
}

// ── String.prototype.toLocaleString ──────────────────────────────────────────

#[test]
fn to_locale_string_is_same_as_to_string_for_strings() {
    // ECMA-262 §22.1.3.28: String.prototype.toLocaleString is implementation-defined
    // but for a simple ASCII string it must return the string itself.
    assert_eq!(
        as_string(&invoke("toLocaleString", vec![s("hello")])),
        "hello"
    );
}

// ── String.prototype.toString (ECMA-262 §22.1.3.29) ─────────────────────────

#[test]
fn to_string_returns_the_string_value() {
    // §22.1.3.29: String.prototype.toString returns the underlying string primitive.
    assert_eq!(as_string(&invoke("toString", vec![s("hello")])), "hello");
}

#[test]
fn to_string_of_empty_string_returns_empty() {
    assert_eq!(as_string(&invoke("toString", vec![s("")])), "");
}

// ── String.prototype.substr (ECMA-262 Annex B §B.2.2.1) ──────────────────────

#[test]
fn substr_extracts_from_start_for_given_length() {
    // Annex B §B.2.2.1: substr(start, length) — length is character count, not end index.
    // "abcdef".substr(1, 3) → "bcd"
    assert_eq!(
        as_string(&invoke(
            "substr",
            vec![s("abcdef"), Value::F64(1.0), Value::F64(3.0)]
        )),
        "bcd"
    );
}

#[test]
fn substr_without_length_extracts_to_end() {
    // "abcdef".substr(2) → "cdef"
    assert_eq!(
        as_string(&invoke("substr", vec![s("abcdef"), Value::F64(2.0)])),
        "cdef"
    );
}

#[test]
fn substr_negative_start_counts_from_end() {
    // "abcdef".substr(-2) → "ef"
    assert_eq!(
        as_string(&invoke("substr", vec![s("abcdef"), Value::F64(-2.0)])),
        "ef"
    );
}

// ── escape / unescape (ECMA-262 Annex B §B.2.1) ─────────────────────────────

#[test]
fn escape_percent_encodes_non_ascii_safe_chars() {
    // Annex B §B.2.1.1: escape leaves A-Z a-z 0-9 @ * _ + - . / alone;
    // encodes space as %20.
    let result = as_string(&invoke("escape", vec![s("hello world")]));
    assert!(result.contains("hello"), "must preserve alpha: {result}");
    assert!(!result.contains(' '), "space must be encoded: {result}");
}

#[test]
fn unescape_reverses_escape_encoding() {
    // Annex B §B.2.1.2: unescape decodes %XX sequences.
    assert_eq!(
        as_string(&invoke("unescape", vec![s("hello%20world")])),
        "hello world"
    );
}

#[allow(dead_code)]
fn _force_object_use(_: Object, _: ObjectKind) {}
