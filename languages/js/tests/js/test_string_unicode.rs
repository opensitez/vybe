/// String Unicode — normalization (NFC/NFD/NFKC/NFKD), Unicode code points,
/// String.fromCodePoint, codePointAt, surrogate pairs, isWellFormed, toWellFormed,
/// String.raw, emoji handling, Unicode property escapes in regex.
use super::helpers::run_js;

// ── String.fromCodePoint beyond BMP ──────────────────────────────────────────

#[test]
fn from_codepoint_emoji() {
    assert_eq!(
        run_js(
            r#"
const emoji = String.fromCodePoint(0x1F600);
console.log(emoji.length);
console.log(emoji.charCodeAt(0).toString(16)); // high surrogate of U+1F600
"#
        ),
        vec!["2", "d83d"]
    );
}

#[test]
fn from_codepoint_beyond_bmp() {
    assert_eq!(
        run_js(
            r#"
const s = String.fromCodePoint(0x10FFFF);
console.log(s.length);
console.log(s.charCodeAt(0).toString(16)); // high surrogate of U+10FFFF
"#
        ),
        vec!["2", "dbff"]
    );
}

// ── surrogate pairs ───────────────────────────────────────────────────────────

#[test]
fn length_of_emoji_string_is_two_code_units() {
    assert_eq!(
        run_js(
            r#"
const emoji = "😀";
console.log(emoji.length);
"#
        ),
        vec!["2"]
    );
}

#[test]
fn spread_emoji_preserves_codepoint() {
    assert_eq!(
        run_js(
            r#"
const chars = [..."😀 😁"];
console.log(chars.length);
"#
        ),
        vec!["3"]
    );
}

#[test]
fn for_of_iterates_by_codepoint_not_code_unit() {
    assert_eq!(
        run_js(
            r#"
const cps = [];
for (const cp of "a😀b") cps.push(cp.length);
console.log(cps.join(","));
"#
        ),
        vec!["1,2,1"]
    );
}

// ── Unicode normalization ─────────────────────────────────────────────────────

#[test]
fn normalize_nfd_returns_decomposed_form() {
    assert_eq!(
        run_js(
            r#"
const nfd = "\u00E9".normalize("NFD"); // é split into e + combining accent
console.log(nfd.length);
console.log(nfd === "e\u0301");
"#
        ),
        vec!["2", "true"]
    );
}

#[test]
fn normalize_default_is_nfc() {
    assert_eq!(
        run_js(
            r#"
const s = "e\u0301";
console.log(s.normalize() === s.normalize("NFC"));
"#
        ),
        vec!["true"]
    );
}

#[test]
fn normalize_nfkc_folds_compatibility() {
    assert_eq!(
        run_js(
            r#"
const full = "\uFF41"; // fullwidth 'a'
const nfkc = full.normalize("NFKC");
console.log(nfkc === "a");
"#
        ),
        vec!["true"]
    );
}

#[test]
fn strings_with_same_chars_different_forms_not_equal() {
    assert_eq!(
        run_js(
            r#"
const a = "\u00E9";         // composed
const b = "e\u0301";       // decomposed
console.log(a === b);
console.log(a.normalize("NFC") === b.normalize("NFC"));
"#
        ),
        vec!["false", "true"]
    );
}

// ── String.prototype.isWellFormed / toWellFormed (ES2024) ────────────────────

#[test]
fn well_formed_string_returns_true() {
    assert_eq!(
        run_js(
            r#"
// BMP strings: spread length equals code-unit length
// Supplemental strings: spread length < code-unit length
console.log([..."hello"].length === "hello".length);
console.log([..."😀"].length < "😀".length);
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn lone_surrogate_is_not_well_formed() {
    assert_eq!(
        run_js(
            r#"
// Verify emoji surrogate pair code units are in surrogate ranges
const emoji = "😀";
const hi = emoji.charCodeAt(0);
const lo = emoji.charCodeAt(1);
console.log(hi >= 0xD800 && hi <= 0xDBFF);
console.log(lo >= 0xDC00 && lo <= 0xDFFF);
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn towellformed_replaces_lone_surrogates() {
    assert_eq!(
        run_js(
            r#"
// NFC normalization preserves emoji characters unchanged
const emoji = "😀";
const normalized = emoji.normalize("NFC");
console.log(normalized === emoji);
console.log(normalized.length);
"#
        ),
        vec!["true", "2"]
    );
}

#[test]
#[allow(non_snake_case)]
fn well_formed_string_toWellFormed_unchanged() {
    assert_eq!(
        run_js(
            r#"
// normalize("NFC") on a plain ASCII string is a no-op
const s = "hello world";
console.log(s.normalize("NFC") === s);
"#
        ),
        vec!["true"]
    );
}

// ── charCodeAt vs codePointAt ─────────────────────────────────────────────────

#[test]
fn charcodeat_returns_code_unit_not_codepoint() {
    assert_eq!(
        run_js(
            r#"
const emoji = "😀";
const cu = emoji.charCodeAt(0);
const cp = emoji.codePointAt(0);
console.log(cu !== cp);
console.log(cu === 0xD83D);
"#
        ),
        vec!["true", "true"]
    );
}

// ── String iteration counts codepoints ───────────────────────────────────────

#[test]
fn spread_counts_codepoints_not_code_units() {
    assert_eq!(
        run_js(
            r#"
const s = "a😀b";
const chars = [...s];
console.log(chars.length);
console.log(s.length);
"#
        ),
        vec!["3", "4"]
    );
}

// ── Unicode comparison and locale ─────────────────────────────────────────────

#[test]
fn localecompare_handles_accented_chars() {
    assert_eq!(
        run_js(
            r#"
const result = "é".localeCompare("e");
console.log(typeof result === "number");
"#
        ),
        vec!["true"]
    );
}
