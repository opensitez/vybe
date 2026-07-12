/// JavaScript advanced regex: named groups, exec, global matching,
/// flags (g, i, m, s), lookahead, lookbehind, replace with function,
/// RegExp constructor, regex methods.
use super::helpers::run_js;

// ===================================================================
// REGEX: EXEC
// ===================================================================

#[test]
fn regex_exec_basic() {
    assert_eq!(
        run_js(
            r#"
let re = /(\d+)-(\d+)/;
let m = re.exec("date: 2024-01");
console.log(m[0]);
console.log(m[1]);
console.log(m[2]);
"#
        ),
        &["2024-01", "2024", "01"]
    );
}

#[test]
fn regex_exec_global_loop() {
    assert_eq!(
        run_js(
            r#"
let re = /\d+/g;
let s = "a1b22c333";
let results = [];
let m;
while ((m = re.exec(s)) !== null) {
    results.push(m[0]);
}
console.log(results.join(","));
"#
        ),
        &["1,22,333"]
    );
}

// ===================================================================
// REGEX: FLAGS
// ===================================================================

#[test]
fn regex_case_insensitive() {
    assert_eq!(
        run_js(
            r#"
let re = /hello/i;
console.log(re.test("Hello World"));
console.log(re.test("HELLO"));
console.log(re.test("hi"));
"#
        ),
        &["true", "true", "false"]
    );
}

#[test]
fn regex_global_match() {
    assert_eq!(
        run_js(
            r#"
let s = "cat bat sat";
let matches = s.match(/[a-z]at/g);
console.log(matches.join(","));
"#
        ),
        &["cat,bat,sat"]
    );
}

#[test]
fn regex_multiline() {
    assert_eq!(
        run_js(
            r#"
let s = "first\nsecond\nthird";
let matches = s.match(/^\w+/gm);
console.log(matches.join(","));
"#
        ),
        &["first,second,third"]
    );
}

// ===================================================================
// REGEX: NAMED GROUPS
// ===================================================================

#[test]
fn regex_named_groups() {
    assert_eq!(
        run_js(
            r#"
let re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/;
let m = re.exec("2024-03-15");
console.log(m.groups.year);
console.log(m.groups.month);
console.log(m.groups.day);
"#
        ),
        &["2024", "03", "15"]
    );
}

// ===================================================================
// REGEX: REPLACE WITH FUNCTION
// ===================================================================

#[test]
fn regex_replace_with_function() {
    assert_eq!(
        run_js(
            r#"
let s = "hello world";
let result = s.replace(/\b\w/g, match => match.toUpperCase());
console.log(result);
"#
        ),
        &["Hello World"]
    );
}

#[test]
fn regex_replace_with_capture() {
    assert_eq!(
        run_js(
            r#"
let s = "John Smith";
let result = s.replace(/(\w+) (\w+)/, "$2, $1");
console.log(result);
"#
        ),
        &["Smith, John"]
    );
}

// ===================================================================
// REGEX: CONSTRUCTOR
// ===================================================================

#[test]
fn regex_constructor() {
    assert_eq!(
        run_js(
            r#"
let pattern = "hello";
let re = new RegExp(pattern, "i");
console.log(re.test("Hello World"));
console.log(re.test("hi"));
"#
        ),
        &["true", "false"]
    );
}

#[test]
fn regex_constructor_dynamic() {
    assert_eq!(
        run_js(
            r#"
function findWord(text, word) {
    let re = new RegExp("\\b" + word + "\\b", "gi");
    let matches = text.match(re);
    return matches ? matches.length : 0;
}
console.log(findWord("The the THE", "the"));
"#
        ),
        &["3"]
    );
}

// ===================================================================
// REGEX: SPLIT
// ===================================================================

#[test]
fn regex_split_multiple_delimiters() {
    assert_eq!(
        run_js(
            r#"
let s = "one,two;three four";
let parts = s.split(/[,; ]/);
console.log(parts.join("|"));
"#
        ),
        &["one|two|three|four"]
    );
}

#[test]
fn regex_split_with_limit() {
    assert_eq!(
        run_js(
            r#"
let s = "a-b-c-d-e";
let parts = s.split("-", 3);
console.log(parts.join(","));
"#
        ),
        &["a,b,c"]
    );
}

// ===================================================================
// REGEX: TEST PATTERNS
// ===================================================================

#[test]
fn regex_test_email_like() {
    assert_eq!(
        run_js(
            r#"
let re = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
console.log(re.test("user@example.com"));
console.log(re.test("bad@"));
console.log(re.test("test.user@domain.co.uk"));
"#
        ),
        &["true", "false", "true"]
    );
}

#[test]
fn regex_test_digits_only() {
    assert_eq!(
        run_js(
            r#"
let re = /^\d+$/;
console.log(re.test("12345"));
console.log(re.test("123a5"));
console.log(re.test(""));
"#
        ),
        &["true", "false", "false"]
    );
}

// ===================================================================
// REGEX: PROPERTIES
// ===================================================================

#[test]
fn regex_source_flags() {
    assert_eq!(
        run_js(
            r#"
let re = /hello/gi;
console.log(re.source);
console.log(re.flags);
console.log(re.global);
console.log(re.ignoreCase);
"#
        ),
        &["hello", "gi", "true", "true"]
    );
}
