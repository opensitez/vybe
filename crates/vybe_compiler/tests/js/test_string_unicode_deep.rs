/// String.prototype.normalize, unicode edge cases, codePoint methods

use super::helpers::run_js;

#[test]
fn string_normalize_nfc() {
    assert_eq!(run_js(r#"
// é as combining: e + combining accent
const composed = "\u00E9"; // é precomposed
const decomposed = "e\u0301"; // e + combining accent
console.log(composed.length);
console.log(decomposed.length);
console.log(composed === decomposed);
console.log(decomposed.normalize("NFC") === composed);
"#), vec!["1", "2", "false", "true"]);
}

#[test]
fn string_normalize_nfd() {
    assert_eq!(run_js(r#"
const composed = "\u00E9";
const nfd = composed.normalize("NFD");
console.log(nfd.length);
console.log(nfd.charCodeAt(0));
console.log(nfd.charCodeAt(1)); // combining acute accent
"#), vec!["2", "101", "769"]);
}

#[test]
fn code_point_at_bmp() {
    assert_eq!(run_js(r#"
const s = "ABC";
console.log(s.codePointAt(0));
console.log(s.codePointAt(1));
"#), vec!["65", "66"]);
}

#[test]
fn code_point_at_surrogate_pair() {
    assert_eq!(run_js(r#"
// 𝌆 is U+1D306, a supplemental character stored as a surrogate pair
const s = "𝌆";
console.log(s.length); // 2 UTF-16 code units
console.log(s.charCodeAt(0).toString(16)); // "d834" — high surrogate
"#), vec!["2", "d834"]);
}

#[test]
fn string_from_code_point() {
    assert_eq!(run_js(r#"
const s = String.fromCodePoint(65, 66, 67);
console.log(s);
"#), vec!["ABC"]);
}

#[test]
fn string_from_code_point_supplemental() {
    assert_eq!(run_js(r#"
const s = String.fromCodePoint(119558); // 𝌆 U+1D306
console.log(s.length); // 2 code units
console.log(s.charCodeAt(0).toString(16)); // "d834" — high surrogate
"#), vec!["2", "d834"]);
}

#[test]
fn for_of_iterates_code_points() {
    assert_eq!(run_js(r#"
const emoji = "😀"; // actual emoji literal (2 code units, 1 codepoint)
const chars = [...emoji];
console.log(chars.length); // 1 character
"#), vec!["1"]);
}

#[test]
fn string_char_at_vs_index() {
    assert_eq!(run_js(r#"
const s = "hello";
console.log(s.charAt(1));
console.log(s[1]);
console.log(s.charAt(99));  // "" for out of bounds
console.log(s[99]);          // undefined for out of bounds
"#), vec!["e", "e", "", "undefined"]);
}

#[test]
fn string_repeat_empty_and_multiple() {
    assert_eq!(run_js(r#"
console.log("abc".repeat(0));
console.log("abc".repeat(1));
console.log("ab".repeat(3));
"#), vec!["", "abc", "ababab"]);
}

#[test]
fn string_split_unicode() {
    assert_eq!(run_js(r#"
const s = "a,b,c";
const parts = s.split(",");
console.log(parts.length);
console.log(parts.join("|"));
"#), vec!["3", "a|b|c"]);
}

#[test]
fn string_includes_works_with_unicode() {
    assert_eq!(run_js(r#"
const s = "Hello, 世界!";
console.log(s.includes("世界"));
console.log(s.includes("World"));
"#), vec!["true", "false"]);
}

#[test]
fn string_slice_with_unicode() {
    assert_eq!(run_js(r#"
const s = "αβγδ";
console.log(s.slice(1, 3));
"#), vec!["βγ"]);
}
