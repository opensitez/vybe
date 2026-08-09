use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: String Inspection (`startsWith`, `endsWith`, `includes`, `search`) & RegEx Traps
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_string_starts_with_basic() {
    let src = r#"
const str = "JavaScript";
console.log(`${str.startsWith("Java")}:${str.startsWith("Script")}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_string_starts_with_position_offset() {
    let src = r#"
const str = "JavaScript";
console.log(str.startsWith("Script", 4));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_ends_with_basic() {
    let src = r#"
const str = "JavaScript";
console.log(`${str.endsWith("Script")}:${str.endsWith("Java")}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_string_ends_with_length_parameter() {
    let src = r#"
const str = "JavaScript";
console.log(str.endsWith("Java", 4));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_includes_basic() {
    let src = r#"
const str = "hello world";
console.log(`${str.includes("world")}:${str.includes("foo")}`);
"#;
    assert_eq!(run_js(src), vec!["true:false"]);
}

#[test]
fn test_js_string_includes_position_offset() {
    let src = r#"
const str = "hello world";
console.log(str.includes("hello", 1));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_string_includes_negative_position_is_zero() {
    let src = r#"
console.log("abc".includes("a", -1));
console.log("abc".includes("b", -1));
"#;
    assert_eq!(run_js(src), vec!["true", "true"]);
}

#[test]
fn test_js_string_starts_with_regex_argument_throws_typeerror() {
    let src = r#"
try {
    "test".startsWith(/test/);
} catch (e) {
    console.log("startsWith RegEx TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["startsWith RegEx TypeError"]);
}

#[test]
fn test_js_string_ends_with_regex_argument_throws_typeerror() {
    let src = r#"
try {
    "test".endsWith(/test/);
} catch (e) {
    console.log("endsWith RegEx TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["endsWith RegEx TypeError"]);
}

#[test]
fn test_js_string_includes_regex_argument_throws_typeerror() {
    let src = r#"
try {
    "test".includes(/test/);
} catch (e) {
    console.log("includes RegEx TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["includes RegEx TypeError"]);
}

#[test]
fn test_js_string_search_regex_basic() {
    let src = r#"
const str = "hello 123 world";
console.log(str.search(/\d+/));
"#;
    assert_eq!(run_js(src), vec!["6"]);
}

#[test]
fn test_js_string_search_number_argument_is_coerced_to_string() {
    let src = r#"
console.log("abc123".search(123));
"#;
    assert_eq!(run_js(src), vec!["3"]);
}

#[test]
fn test_js_string_search_no_match_returns_minus_one() {
    let src = r#"
const str = "hello world";
console.log(str.search(/xyz/));
"#;
    assert_eq!(run_js(src), vec!["-1"]);
}

#[test]
fn test_js_string_search_string_argument_coerced_to_regex() {
    let src = r#"
const str = "hello 123 world";
console.log(str.search("123")); // String argument in search() is implicitly converted to new RegExp("123")!
"#;
    assert_eq!(run_js(src), vec!["6"]);
}

#[test]
fn test_js_string_search_symbol_search_protocol() {
    let src = r#"
const customSearcher = {
    [Symbol.search](target) {
        return target.indexOf("custom");
    }
};
console.log("prefix_custom_suffix".search(customSearcher));
"#;
    assert_eq!(run_js(src), vec!["7"]);
}

#[test]
fn test_js_string_starts_with_custom_match_symbol_override() {
    let src = r#"
const fakeRegex = {
    [Symbol.match]: false, // Setting Symbol.match to false allows object with RegExp prototype to pass startsWith!
    toString() { return "abc"; }
};
console.log("abcdef".startsWith(fakeRegex));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_starts_with_empty_search_string() {
    let src = r#"
console.log("abc".startsWith(""));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_ends_with_empty_search_string() {
    let src = r#"
console.log("abc".endsWith(""));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_includes_empty_search_string() {
    let src = r#"
console.log("abc".includes(""));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_starts_with_case_sensitive() {
    let src = r#"
console.log("abc".startsWith("A"));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_string_search_ignores_global_g_flag() {
    let src = r#"
const re = /a/g;
re.lastIndex = 5;
const pos = "cat bat".search(re);
console.log(pos + "|lastIndex=" + re.lastIndex); // search ignores g flag and does not mutate lastIndex!
"#;
    assert_eq!(run_js(src), vec!["1|lastIndex=5"]);
}

#[test]
fn test_js_string_inspection_null_or_undefined_this_throws_typeerror() {
    let src = r#"
try {
    String.prototype.includes.call(null, "test");
} catch (e) {
    console.log("includes Null This TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["includes Null This TypeError"]);
}

#[test]
fn test_js_string_boundary_positions_and_empty_search_behavior() {
    let src = r#"
console.log(`${"abc".startsWith("a", 10)}:${"abc".endsWith("a", 1)}:${"abc".includes("z", 10)}:${"".includes("")}:${"abc".includes("")}`);
"#;
    assert_eq!(run_js(src), vec!["false:true:false:true:true"]);
}

#[test]
fn test_js_string_starts_with_nan_position_treats_as_zero() {
    let src = r#"
console.log("abc".startsWith("a", NaN));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
