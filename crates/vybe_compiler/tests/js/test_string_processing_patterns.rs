/// String processing patterns — template engines, parsers, formatters
use super::helpers::run_js;

#[test]
fn string_interpolation_template() {
    assert_eq!(
        run_js(
            r#"
function interpolate(template, data) {
    return template.replace(/\{\{(\w+)\}\}/g, (_, key) => data[key] ?? "");
}
const result = interpolate("Hello {{name}}, you have {{count}} messages!", {
    name: "Alice",
    count: 5
});
console.log(result);
"#
        ),
        vec!["Hello Alice, you have 5 messages!"]
    );
}

#[test]
fn camel_to_snake_case() {
    assert_eq!(
        run_js(
            r#"
function toSnakeCase(str) {
    return str.replace(/([A-Z])/g, c => "_" + c.toLowerCase()).replace(/^_/, "");
}
console.log(toSnakeCase("helloWorld"));
console.log(toSnakeCase("camelCaseString"));
console.log(toSnakeCase("simpleTest"));
"#
        ),
        vec!["hello_world", "camel_case_string", "simple_test"]
    );
}

#[test]
fn snake_to_camel_case() {
    assert_eq!(
        run_js(
            r#"
function toCamelCase(str) {
    return str.replace(/_(\w)/g, (_, c) => c.toUpperCase());
}
console.log(toCamelCase("hello_world"));
console.log(toCamelCase("some_variable_name"));
"#
        ),
        vec!["helloWorld", "someVariableName"]
    );
}

#[test]
fn word_count() {
    assert_eq!(
        run_js(
            r#"
function wordCount(str) {
    return str.trim().split(/\s+/).filter(Boolean).length;
}
console.log(wordCount("hello world foo"));
console.log(wordCount("  spaced  words  "));
console.log(wordCount(""));
"#
        ),
        vec!["3", "2", "0"]
    );
}

#[test]
fn truncate_with_ellipsis() {
    assert_eq!(
        run_js(
            r#"
function truncate(str, max, ellipsis = "...") {
    if (str.length <= max) return str;
    return str.slice(0, max - ellipsis.length) + ellipsis;
}
console.log(truncate("Hello, World!", 8));
console.log(truncate("Short", 10));
"#
        ),
        vec!["Hello...", "Short"]
    );
}

#[test]
fn parse_csv_line() {
    assert_eq!(
        run_js(
            r#"
function parseCSV(line) {
    return line.split(",").map(s => s.trim());
}
const row = parseCSV("Alice, 30, Engineer");
console.log(row[0]);
console.log(row[1]);
console.log(row[2]);
"#
        ),
        vec!["Alice", "30", "Engineer"]
    );
}

#[test]
fn string_is_palindrome() {
    assert_eq!(
        run_js(
            r#"
function isPalindrome(str) {
    const clean = str.toLowerCase().replace(/[^a-z0-9]/g, "");
    return clean === [...clean].reverse().join("");
}
console.log(isPalindrome("racecar"));
console.log(isPalindrome("A man a plan a canal Panama"));
console.log(isPalindrome("hello"));
"#
        ),
        vec!["true", "true", "false"]
    );
}

#[test]
fn count_occurrences() {
    assert_eq!(
        run_js(
            r#"
function count(str, sub) {
    let n = 0, pos = 0;
    while ((pos = str.indexOf(sub, pos)) !== -1) { n++; pos += sub.length; }
    return n;
}
console.log(count("abcabcabc", "abc"));
console.log(count("hello world hello", "hello"));
console.log(count("aaa", "aa")); // non-overlapping
"#
        ),
        vec!["3", "2", "1"]
    );
}

#[test]
fn format_number_with_commas() {
    assert_eq!(
        run_js(
            r#"
function formatNumber(n) {
    return n.toLocaleString("en-US");
}
// Or use regex:
function formatNumberRegex(n) {
    return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}
console.log(formatNumberRegex(1234567));
console.log(formatNumberRegex(1000));
console.log(formatNumberRegex(42));
"#
        ),
        vec!["1,234,567", "1,000", "42"]
    );
}

#[test]
fn string_lines_split_and_process() {
    assert_eq!(
        run_js(
            r#"
const text = "line1\nline2\nline3\n";
const lines = text.split("\n").filter(Boolean);
console.log(lines.length);
console.log(lines[0]);
console.log(lines[2]);
"#
        ),
        vec!["3", "line1", "line3"]
    );
}
