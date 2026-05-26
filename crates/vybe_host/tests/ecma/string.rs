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

use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{Chunk, Op, VM};
use vybe_host::{Capabilities, register_with_capabilities};

fn invoke(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<ecma-string-test>");
    let import_idx = chunk.add_import("ecma:string", name);
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

// ── length / charAt / charCodeAt / codePointAt ────────────────────

#[test]
fn length_counts_utf16_code_units() {
    assert_eq!(invoke("length", vec![s("hello")]), Value::F64(5.0));
    assert_eq!(invoke("length", vec![s("")]), Value::F64(0.0));
}

#[test]
fn char_at_returns_single_character() {
    assert_eq!(as_string(&invoke("charAt", vec![s("hello"), Value::F64(1.0)])), "e");
}

#[test]
fn char_at_out_of_range_returns_empty_string() {
    // ECMA-262: out-of-range returns empty string (NOT undefined).
    assert_eq!(as_string(&invoke("charAt", vec![s("hi"), Value::F64(99.0)])), "");
}

#[test]
fn char_code_at_returns_unit_code() {
    assert_eq!(invoke("charCodeAt", vec![s("Aa"), Value::F64(0.0)]), Value::F64(65.0));
    assert_eq!(invoke("charCodeAt", vec![s("Aa"), Value::F64(1.0)]), Value::F64(97.0));
}

#[test]
fn at_supports_negative_index() {
    // ECMA-262 §22.1.3.1 (added in ES2022).
    assert_eq!(as_string(&invoke("at", vec![s("hello"), Value::F64(-1.0)])), "o");
    assert_eq!(as_string(&invoke("at", vec![s("hello"), Value::F64(0.0)])), "h");
}

// ── concat ────────────────────────────────────────────────────────

#[test]
fn concat_joins_arguments() {
    assert_eq!(as_string(&invoke("concat", vec![s("ab"), s("cd"), s("ef")])), "abcdef");
}

// ── substring / slice ─────────────────────────────────────────────

#[test]
fn substring_extracts_range() {
    // substring(start, end) — both clamped non-negative.
    assert_eq!(as_string(&invoke("substring", vec![s("abcdef"), Value::F64(1.0), Value::F64(4.0)])), "bcd");
}

#[test]
fn substring_swaps_when_start_greater_than_end() {
    // ECMA-262: substring swaps args if start > end.
    assert_eq!(as_string(&invoke("substring", vec![s("abcdef"), Value::F64(4.0), Value::F64(1.0)])), "bcd");
}

#[test]
fn slice_extracts_range() {
    assert_eq!(as_string(&invoke("slice", vec![s("abcdef"), Value::F64(1.0), Value::F64(4.0)])), "bcd");
}

#[test]
fn slice_supports_negative_indices() {
    // slice — negative indices count from end (substring does not).
    assert_eq!(as_string(&invoke("slice", vec![s("abcdef"), Value::F64(-3.0)])), "def");
    assert_eq!(as_string(&invoke("slice", vec![s("abcdef"), Value::F64(-3.0), Value::F64(-1.0)])), "de");
}

// ── includes / indexOf / lastIndexOf ──────────────────────────────

#[test]
fn includes_finds_substring() {
    assert_eq!(invoke("includes", vec![s("hello world"), s("world")]), Value::Bool(true));
    assert_eq!(invoke("includes", vec![s("hello world"), s("XYZ")]), Value::Bool(false));
}

#[test]
fn includes_with_start_position() {
    assert_eq!(invoke("includes", vec![s("hello world"), s("hello"), Value::F64(1.0)]), Value::Bool(false));
}

#[test]
fn index_of_returns_position_or_minus_one() {
    assert_eq!(invoke("indexOf", vec![s("hello world"), s("world")]), Value::F64(6.0));
    assert_eq!(invoke("indexOf", vec![s("hello world"), s("XYZ")]), Value::F64(-1.0));
}

#[test]
fn last_index_of_finds_last_occurrence() {
    assert_eq!(invoke("lastIndexOf", vec![s("ababab"), s("ab")]), Value::F64(4.0));
    assert_eq!(invoke("lastIndexOf", vec![s("ababab"), s("XYZ")]), Value::F64(-1.0));
}

// ── startsWith / endsWith ─────────────────────────────────────────

#[test]
fn starts_with_returns_bool() {
    assert_eq!(invoke("startsWith", vec![s("hello"), s("hel")]), Value::Bool(true));
    assert_eq!(invoke("startsWith", vec![s("hello"), s("ell")]), Value::Bool(false));
}

#[test]
fn ends_with_returns_bool() {
    assert_eq!(invoke("endsWith", vec![s("hello"), s("llo")]), Value::Bool(true));
    assert_eq!(invoke("endsWith", vec![s("hello"), s("hel")]), Value::Bool(false));
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
    assert_eq!(as_string(&invoke("trimStart", vec![s("  hello  ")])), "hello  ");
}

#[test]
fn trim_end_strips_only_trailing() {
    assert_eq!(as_string(&invoke("trimEnd", vec![s("  hello  ")])), "  hello");
}

// ── padStart / padEnd ─────────────────────────────────────────────

#[test]
fn pad_start_pads_to_target_length() {
    assert_eq!(as_string(&invoke("padStart", vec![s("5"), Value::F64(3.0), s("0")])), "005");
}

#[test]
fn pad_start_default_pad_is_space() {
    assert_eq!(as_string(&invoke("padStart", vec![s("hi"), Value::F64(5.0)])), "   hi");
}

#[test]
fn pad_end_pads_to_target_length() {
    assert_eq!(as_string(&invoke("padEnd", vec![s("5"), Value::F64(3.0), s(".")])), "5..");
}

// ── repeat ────────────────────────────────────────────────────────

#[test]
fn repeat_concatenates_n_times() {
    assert_eq!(as_string(&invoke("repeat", vec![s("ab"), Value::F64(3.0)])), "ababab");
}

#[test]
fn repeat_zero_returns_empty() {
    assert_eq!(as_string(&invoke("repeat", vec![s("ab"), Value::F64(0.0)])), "");
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
    assert_eq!(array_strings(&invoke("split", vec![s("a,b,c"), s(",")])), vec!["a", "b", "c"]);
}

#[test]
fn split_with_limit_truncates() {
    assert_eq!(
        array_strings(&invoke("split", vec![s("a,b,c,d"), s(","), Value::F64(2.0)])),
        vec!["a", "b"]
    );
}

#[test]
fn split_empty_separator_yields_per_character() {
    assert_eq!(array_strings(&invoke("split", vec![s("abc"), s("")])), vec!["a", "b", "c"]);
}

// ── String.fromCharCode / fromCodePoint (constructor statics) ─────

#[test]
fn from_char_code_builds_string() {
    assert_eq!(as_string(&invoke("fromCharCode", vec![Value::F64(65.0), Value::F64(97.0)])), "Aa");
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
    assert_eq!(invoke("localeCompare", vec![s("abc"), s("abc")]), Value::I32(0));
}

#[test]
fn locale_compare_returns_positive_for_greater() {
    if let Value::I32(n) = invoke("localeCompare", vec![s("z"), s("a")]) {
        assert!(n > 0);
    } else {
        panic!("localeCompare should return I32");
    }
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
        as_string(&invoke("encodeURI", vec![s("https://example.com/path?q=1")])),
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

#[allow(dead_code)]
fn _force_object_use(_: Object, _: ObjectKind) {}
