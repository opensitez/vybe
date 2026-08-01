use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: String Methods & Unicode — split, rsplit, partition, rpartition, replace, strip, casefold, isidentifier
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_string_split_and_rsplit() {
    let src = r#"
text = "one,two,three,four,five"
print(text.split(",", 2))
print(text.rsplit(",", 2))
whitespace = "  a  b   c  "
print(whitespace.split())
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['one', 'two', 'three,four,five']",
            "['one,two,three', 'four', 'five']",
            "['a', 'b', 'c']"
        ]
    );
}

#[test]
fn test_py_string_partition_and_rpartition() {
    let src = r#"
url = "https://example.com/path/to/page"
print(url.partition("://"))
print(url.rpartition("/"))
print("nohost".partition("://"))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "('https', '://', 'example.com/path/to/page')",
            "('https://example.com/path/to', '/', 'page')",
            "('nohost', '', '')"
        ]
    );
}

#[test]
fn test_py_string_casefold_unicode_case_insensitive() {
    let src = r#"
# German sharp s 'ß' casefolds to 'ss'
s1 = "STRASSE"
s2 = "straße"
print(s1.lower() == s2.lower())
print(s1.casefold() == s2.casefold())
"#;
    assert_eq!(run_python(src), vec!["False", "True"]);
}

#[test]
fn test_py_string_isidentifier() {
    let src = r#"
print("var_name".isidentifier())
print("_private".isidentifier())
print("123var".isidentifier())
print("class".isidentifier())  # identifier keyword check
"#;
    assert_eq!(run_python(src), vec!["True", "True", "False", "True"]);
}

#[test]
fn test_py_string_strip_lstrip_rstrip_chars() {
    let src = r#"
text = "...hello world!!!"
print(text.strip(".!"))
print(text.lstrip("."))
print(text.rstrip("!"))
"#;
    assert_eq!(
        run_python(src),
        vec!["hello world", "hello world!!!", "...hello world"]
    );
}

#[test]
fn test_py_string_removeprefix_removesuffix_py39() {
    let src = r#"
filename = "document.pdf.txt"
print(filename.removeprefix("document."))
print(filename.removesuffix(".txt"))
print(filename.removeprefix("missing"))
"#;
    assert_eq!(
        run_python(src),
        vec!["pdf.txt", "document.pdf", "document.pdf.txt"]
    );
}

#[test]
fn test_py_string_replace_count_limit() {
    let src = r#"
text = "foo bar foo baz foo"
print(text.replace("foo", "qux"))
print(text.replace("foo", "qux", 2))
"#;
    assert_eq!(
        run_python(src),
        vec!["qux bar qux baz qux", "qux bar qux baz foo"]
    );
}

#[test]
fn test_py_string_find_index_rfind_rindex() {
    let src = r#"
text = "banana"
print(text.find("an"))
print(text.rfind("an"))
print(text.index("ba"))
try:
    text.index("xyz")
except ValueError:
    print("ValueError")
"#;
    assert_eq!(run_python(src), vec!["1", "3", "0", "ValueError"]);
}

#[test]
fn test_py_string_center_ljust_rjust_zfill() {
    let src = r#"
s = "42"
print(s.zfill(5))
print(s.center(6, "-"))
print(s.ljust(5, "."))
print(s.rjust(5, "."))
"#;
    assert_eq!(run_python(src), vec!["00042", "--42--", "42...", "...42"]);
}

#[test]
fn test_py_string_predicate_checks() {
    let src = r#"
print("12345".isdigit())
print("12345".isdecimal())
print("½".isnumeric())
print("½".isdigit())
print("Hello World".istitle())
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "True", "True", "False", "True"]
    );
}

// ── Search methods: code-point indices and the start/end window ────────
//
// Python string indices count CODE POINTS (`unifiedstringplan.md` Axis 1,
// unit `scalar`), so `find`/`index`/`rfind`/`rindex` must agree with `len`,
// `s[i]` and `s[a:b]`. They were bound to the UTF-16 helpers, which answered
// 3 for `"a😀b".find("b")`.
//
// The receiver goes through a function in the astral case so the value is not
// constant-foldable — a literal receiver can be resolved at compile time and
// never reach the emitter.

#[test]
fn test_py_search_methods_count_code_points() {
    let src = r#"
def f(x): return x
s = f("a\U0001F600b\U0001F600c")
print(s.index("b"), s.find("b"), s.rfind("c"), s.rindex("c"))
print(len(s), s[1], s[1:3])
"#;
    assert_eq!(run_python(src), vec!["2 2 4 4", "5 😀 😀b"]);
}

/// `ord` is a CODE POINT. `charCodeAt(s, 0)` answers 55357 for `😀` — the
/// leading surrogate — which is not a character Python can name.
#[test]
fn test_py_ord_is_a_code_point() {
    let src = r#"
print(ord("a"), ord("é"), ord("\U0001F600"))
"#;
    assert_eq!(run_python(src), vec!["97 233 128512"]);
}

/// `find`/`rfind` take an optional `start` and `end`, and the answer stays in
/// whole-string coordinates. The raw binding ignored the extra arguments and
/// left them on the stack, so `"abcabc".find("b", 2)` produced stack garbage
/// rather than 4 — on pure ASCII, with no test covering it.
#[test]
fn test_py_find_start_and_end_window() {
    let src = r#"
def f(x): return x
t = f("abcabc")
print(t.find("b"), t.find("b", 2), t.find("z"), t.rfind("b"))
print(t.find("b", 2, 3), t.find("c", -2), t.rfind("b", 0, 2))
try:
    t.rindex("z")
except ValueError:
    print("ValueError")
"#;
    assert_eq!(
        run_python(src),
        vec!["1 4 -1 4", "-1 5 1", "ValueError"]
    );
}
