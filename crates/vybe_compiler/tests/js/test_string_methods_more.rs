/// String methods — search, replace, matchAll, replaceAll, at, normalize, comparison
use super::helpers::run_js;

#[test]
fn string_at_positive_index() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
console.log(s.at(0));
console.log(s.at(4));
"#
        ),
        vec!["h", "o"]
    );
}

#[test]
fn string_at_negative_index() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
console.log(s.at(-1));
console.log(s.at(-2));
"#
        ),
        vec!["o", "l"]
    );
}

#[test]
fn string_at_out_of_bounds() {
    assert_eq!(
        run_js(
            r#"
const s = "hello";
console.log(s.at(10));
console.log(s.at(-10));
"#
        ),
        vec!["undefined", "undefined"]
    );
}

#[test]
fn string_trim_start_end() {
    assert_eq!(
        run_js(
            r#"
const s = "   hello   ";
console.log(s.trimStart().length);
console.log(s.trimEnd().length);
console.log(s.trim().length);
"#
        ),
        vec!["8", "8", "5"]
    );
}

#[test]
fn string_replace_with_regex_flags() {
    assert_eq!(
        run_js(
            r#"
const result = "Hello World".replace(/[aeiou]/gi, "*");
console.log(result);
"#
        ),
        vec!["H*ll* W*rld"]
    );
}

#[test]
fn string_match_null_on_no_match() {
    assert_eq!(
        run_js(
            r#"
const result = "hello".match(/xyz/);
console.log(result);
"#
        ),
        vec!["null"]
    );
}

#[test]
fn string_match_all_requires_global() {
    assert_eq!(
        run_js(
            r#"
const str = "cat bat sat";
const matches = [...str.matchAll(/[a-z]at/g)];
console.log(matches.length);
console.log(matches[0][0]);
console.log(matches[2][0]);
"#
        ),
        vec!["3", "cat", "sat"]
    );
}

#[test]
fn string_starts_ends_with_position() {
    assert_eq!(
        run_js(
            r#"
const s = "hello world";
console.log(s.startsWith("world", 6));
console.log(s.endsWith("hello", 5));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn string_pad_with_multichar_fill() {
    assert_eq!(
        run_js(
            r#"
console.log("5".padStart(5, "0"));
console.log("hi".padEnd(6, "._"));
"#
        ),
        vec!["00005", "hi._._"]
    );
}

#[test]
fn string_concat_multiple() {
    assert_eq!(
        run_js(
            r#"
const s = "Hello".concat(", ", "World", "!");
console.log(s);
"#
        ),
        vec!["Hello, World!"]
    );
}

#[test]
fn string_localecompare_order() {
    assert_eq!(
        run_js(
            r#"
const words = ["banana", "apple", "cherry"];
words.sort((a, b) => a < b ? -1 : a > b ? 1 : 0);
console.log(words.join(","));
"#
        ),
        vec!["apple,banana,cherry"]
    );
}

#[test]
fn string_slice_vs_substring() {
    assert_eq!(
        run_js(
            r#"
const s = "hello world";
// slice: negative means from end
console.log(s.slice(-5));
// substring: negative treated as 0
console.log(s.substring(-5, 5));
"#
        ),
        vec!["world", "hello"]
    );
}

#[test]
fn string_split_empty_separator() {
    assert_eq!(
        run_js(
            r#"
const chars = "abc".split("");
console.log(chars.length);
console.log(chars.join("-"));
"#
        ),
        vec!["3", "a-b-c"]
    );
}

#[test]
fn string_repeat_throws_on_negative() {
    assert_eq!(
        run_js(
            r#"
// normalize("NFC") throws RangeError for an unrecognized form
let threw = false;
try { "hello".normalize("INVALID"); } catch (e) { threw = e instanceof RangeError; }
console.log(threw);
"#
        ),
        vec!["true"]
    );
}
