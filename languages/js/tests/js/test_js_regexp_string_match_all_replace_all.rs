use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: String Matching & Replacement (`matchAll`, `replaceAll`, `Symbol.matchAll`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_string_matchall_returns_regex_iterator() {
    let src = r#"
const re = /t(e)(st(\d?))/g;
const str = "test1test2";
const matches = [...str.matchAll(re)];
console.log(matches.length + "|" + matches[0][0] + "|" + matches[1][0]);
"#;
    assert_eq!(run_js(src), vec!["2|test1|test2"]);
}

#[test]
fn test_js_string_matchall_requires_global_regexp_throws_typeerror() {
    let src = r#"
try {
    "hello".matchAll(/l/); // Non-global RegExp throws TypeError!
} catch (e) {
    console.log("matchAll Non-Global TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["matchAll Non-Global TypeError"]);
}

#[test]
fn test_js_string_matchall_with_string_argument_coerced_to_global_regexp() {
    let src = r#"
const matches = [..."hello world".matchAll("l")];
console.log(matches.length + "|" + matches.map(m => m.index).join(","));
"#;
    assert_eq!(run_js(src), vec!["3|2,3,9"]);
}

#[test]
fn test_js_string_replaceall_basic_string_replacement() {
    let src = r#"
const str = "foo-bar-foo";
console.log(str.replaceAll("foo", "baz"));
"#;
    assert_eq!(run_js(src), vec!["baz-bar-baz"]);
}

#[test]
fn test_js_string_replaceall_with_global_regexp() {
    let src = r#"
const str = "apple banana apple";
console.log(str.replaceAll(/apple/g, "orange"));
"#;
    assert_eq!(run_js(src), vec!["orange banana orange"]);
}

#[test]
fn test_js_string_replaceall_non_global_regexp_throws_typeerror() {
    let src = r#"
try {
    "test".replaceAll(/t/, "x");
} catch (e) {
    console.log("replaceAll Non-Global TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["replaceAll Non-Global TypeError"]);
}

#[test]
fn test_js_string_replaceall_replacement_function() {
    let src = r#"
const str = "1 2 3 4";
const res = str.replaceAll(/\d/g, match => String(Number(match) * 10));
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["10 20 30 40"]);
}

#[test]
fn test_js_string_matchall_capturing_groups() {
    let src = r#"
const str = "a1 b2 c3";
const matches = [...str.matchAll(/([a-z])(\d)/g)];
console.log(matches.map(m => `${m[1]}:${m[2]}`).join(","));
"#;
    assert_eq!(run_js(src), vec!["a:1,b:2,c:3"]);
}

#[test]
fn test_js_string_matchall_custom_symbol_matchall_matcher() {
    let src = r#"
const customMatcher = {
    [Symbol.matchAll](string) {
        return ["Custom1", "Custom2"][Symbol.iterator]();
    }
};
console.log([..."input".matchAll(customMatcher)].join(","));
"#;
    assert_eq!(run_js(src), vec!["Custom1,Custom2"]);
}

#[test]
fn test_js_string_replaceall_special_replacement_patterns() {
    let src = r#"
const str = "world";
console.log(str.replaceAll("world", "Hello $$ $` $& $'"));
"#;
    assert_eq!(run_js(src), vec!["Hello $  world "]);
}

#[test]
fn test_js_string_replaceall_empty_search_string() {
    let src = r#"
const str = "abc";
console.log(str.replaceAll("", "_"));
"#;
    assert_eq!(run_js(src), vec!["_a_b_c_"]);
}

#[test]
fn test_js_string_matchall_empty_regex_global() {
    let src = r#"
const matches = [..."hi".matchAll(/(?:)/g)];
console.log(matches.length + "|" + matches.map(m => m.index).join(","));
"#;
    assert_eq!(run_js(src), vec!["3|0,1,2"]);
}

#[test]
fn test_js_string_replaceall_with_named_capture_groups() {
    let src = r#"
const str = "2026-07-22";
const re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/g;
const res = str.replaceAll(re, "$<month>/$<day>/$<year>");
console.log(res);
"#;
    assert_eq!(run_js(src), vec!["07/22/2026"]);
}

#[test]
fn test_js_string_matchall_iterator_protocol_done_state() {
    let src = r#"
const iter = "x".matchAll(/x/g);
const s1 = iter.next();
const s2 = iter.next();
console.log(`${s1.value[0]}|done=${s1.done}`);
console.log(`${s2.value}|done=${s2.done}`);
"#;
    assert_eq!(run_js(src), vec!["x|done=false", "undefined|done=true"]);
}

#[test]
fn test_js_string_replaceall_function_replacer_arguments() {
    let src = r#"
const str = "cat";
str.replaceAll(/(a)/g, (match, p1, offset, string) => {
    console.log(`${match}:${p1}:${offset}:${string}`);
    return "X";
});
"#;
    assert_eq!(run_js(src), vec!["a:a:1:cat"]);
}

#[test]
fn test_js_string_matchall_clones_regexp_lastindex() {
    let src = r#"
const re = /a/g;
re.lastIndex = 2; // Original regex lastIndex should NOT affect matchAll or be modified by it!
const matches = [..."aaa".matchAll(re)];
console.log(matches.length + "|originalLastIndex=" + re.lastIndex);
"#;
    assert_eq!(run_js(src), vec!["3|originalLastIndex=2"]);
}

#[test]
fn test_js_string_replaceall_null_and_undefined_coercion() {
    let src = r#"
console.log("null value undefined".replaceAll("null", "1").replaceAll("undefined", "2"));
"#;
    assert_eq!(run_js(src), vec!["1 value 2"]);
}

#[test]
fn test_js_string_matchall_unicode_flag_surrogate_pairs() {
    let src = r#"
const matches = [..."😀😃".matchAll(/\p{Emoji}/gu)];
console.log(matches.length + "|" + matches[0][0]);
"#;
    assert_eq!(run_js(src), vec!["2|😀"]);
}

#[test]
fn test_js_string_replaceall_symbol_replace_protocol_override() {
    let src = r#"
const customReplacer = {
    [Symbol.replace](string, replacement) {
        return `Overridden:${replacement}`;
    }
};
console.log("input".replaceAll(customReplacer, "ReplacementText"));
"#;
    assert_eq!(run_js(src), vec!["Overridden:ReplacementText"]);
}

#[test]
fn test_js_string_matchall_returns_new_iterator_each_call() {
    let src = r#"
const str = "abc";
const i1 = str.matchAll(/./g);
const i2 = str.matchAll(/./g);
console.log(i1 !== i2);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
