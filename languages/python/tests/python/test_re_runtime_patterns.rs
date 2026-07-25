use crate::helpers::{run_print, run_python_one};

#[test]
fn re_search_finds_substring() {
    assert_eq!(
        run_python_one("import re\nm = re.search('world', 'hello world')\nprint(bool(m))\n"),
        "True"
    );
}

#[test]
fn re_search_no_match_returns_none() {
    assert_eq!(
        run_python_one("import re\nprint(re.search('z', 'abc'))\n"),
        "None"
    );
}

#[test]
fn re_match_at_start() {
    assert_eq!(
        run_python_one("import re\nm = re.match('he', 'hello')\nprint(bool(m))\n"),
        "True"
    );
}

#[test]
fn re_match_not_at_start_fails() {
    assert_eq!(
        run_python_one("import re\nprint(re.match('lo', 'hello'))\n"),
        "None"
    );
}

#[test]
fn re_findall_digits() {
    assert_eq!(
        run_print("import re; re.findall(r'\\d+', 'a1b22c')"),
        "['1', '22']"
    );
}

#[test]
fn re_findall_words() {
    assert_eq!(
        run_print("import re; re.findall(r'[a-z]+', 'Hello World')"),
        "['ello', 'orld']"
    );
}

#[test]
fn re_sub_replace_once() {
    assert_eq!(run_print("import re; re.sub('a', 'x', 'aba')"), "xbx");
}

#[test]
fn re_sub_count_limit() {
    assert_eq!(
        run_print("import re; re.sub('a', 'x', 'aaa', count=2)"),
        "xxa"
    );
}

#[test]
fn re_split_on_whitespace() {
    assert_eq!(
        run_print("import re; re.split(r'\\s+', 'a  b   c')"),
        "['a', 'b', 'c']"
    );
}

#[test]
fn re_split_maxsplit() {
    assert_eq!(
        run_print("import re; re.split(',', 'a,b,c', maxsplit=1)"),
        "['a', 'b,c']"
    );
}

#[test]
fn re_search_group_zero() {
    assert_eq!(
        run_python_one("import re\nm = re.search('(ab)+', 'xxababc')\nprint(m.group(0))\n"),
        "ab"
    );
}

#[test]
fn re_search_group_one() {
    assert_eq!(
        run_python_one("import re\nm = re.search('a(b)c', 'abc')\nprint(m.group(1))\n"),
        "b"
    );
}

#[test]
fn re_compile_reuse_pattern() {
    assert_eq!(
        run_python_one("import re\np = re.compile('[0-9]+')\nprint(p.findall('a1b2'))\n"),
        "['1', '2']"
    );
}

#[test]
fn re_search_case_sensitive_default() {
    assert_eq!(
        run_python_one("import re\nprint(re.search('Hello', 'hello') is None)\n"),
        "True"
    );
}

#[test]
fn re_findall_empty_pattern_matches_empty() {
    assert_eq!(
        run_print("import re; re.findall('', 'abc')"),
        "['', '', '', '']"
    );
}

#[test]
fn re_sub_function_replacement() {
    assert_eq!(
        run_python_one(
            "import re\nprint(re.sub(r'\\d+', lambda m: str(int(m.group(0)) * 2), 'a3b'))\n"
        ),
        "a6b"
    );
}

#[test]
fn re_split_capturing_groups_kept() {
    assert_eq!(
        run_print("import re; re.split('(/)', '/a/b')"),
        "['', '/', 'a', '/', 'b']"
    );
}

#[test]
fn re_search_dot_matches_any_char() {
    assert_eq!(
        run_python_one("import re\nm = re.search('h.l', 'hello')\nprint(m.group(0))\n"),
        "hel"
    );
}

#[test]
fn re_findall_anchored_start() {
    assert_eq!(run_print("import re; re.findall('^a', 'ab ac')"), "['a']");
}

#[test]
fn re_sub_no_match_unchanged() {
    assert_eq!(run_print("import re; re.sub('z', 'Z', 'abc')"), "abc");
}

#[test]
fn re_search_end_position() {
    assert_eq!(
        run_python_one("import re\nm = re.search('world', 'hello world')\nprint(m.end())\n"),
        "11"
    );
}

#[test]
fn re_search_start_position() {
    assert_eq!(
        run_python_one("import re\nm = re.search('world', 'hello world')\nprint(m.start())\n"),
        "6"
    );
}

#[test]
fn re_findall_alternation() {
    assert_eq!(
        run_print("import re; re.findall('cat|dog', 'cat dog bird')"),
        "['cat', 'dog']"
    );
}

#[test]
fn re_split_on_comma() {
    assert_eq!(
        run_print("import re; re.split(',', 'x,y,z')"),
        "['x', 'y', 'z']"
    );
}

#[test]
fn re_match_full_string_optional_end() {
    assert_eq!(
        run_python_one("import re\nm = re.match('hello', 'hello')\nprint(m.group(0))\n"),
        "hello"
    );
}

#[test]
fn re_search_quantifier_plus() {
    assert_eq!(
        run_python_one("import re\nm = re.search('a+', 'baaab')\nprint(m.group(0))\n"),
        "aaa"
    );
}

#[test]
fn re_findall_word_boundaries_simple() {
    assert_eq!(
        run_print("import re; re.findall(r'\\b\\w+\\b', 'hi there')"),
        "['hi', 'there']"
    );
}

#[test]
fn re_sub_backreference_style() {
    assert_eq!(
        run_print("import re; re.sub(r'(a)(b)', r'\\2\\1', 'ab')"),
        "ba"
    );
}

#[test]
fn re_search_none_bool_false() {
    assert_eq!(
        run_python_one("import re\nprint(bool(re.search('z', 'abc')))\n"),
        "False"
    );
}

#[test]
fn re_findall_on_empty_string() {
    assert_eq!(run_print("import re; re.findall('x', '')"), "[]");
}

#[test]
fn re_split_trailing_empty() {
    assert_eq!(
        run_print("import re; re.split(',', 'a,b,')"),
        "['a', 'b', '']"
    );
}

#[test]
fn re_compile_flags_ignorecase() {
    assert_eq!(
        run_python_one("import re\np = re.compile('abc', re.I)\nprint(bool(p.search('AbC')))\n"),
        "True"
    );
}

#[test]
fn re_search_multiline_caret() {
    assert_eq!(
        run_python_one("import re\nm = re.search('^b', 'a\\nb', re.M)\nprint(bool(m))\n"),
        "True"
    );
}

#[test]
fn re_findall_character_class() {
    assert_eq!(
        run_print("import re; re.findall('[aeiou]', 'hello')"),
        "['e', 'o']"
    );
}

#[test]
fn re_sub_numeric_backref_in_pattern() {
    assert_eq!(
        run_print("import re; re.sub(r'(a)(\\d)', r'\\2\\1', 'a1')"),
        "1a"
    );
}

#[test]
fn re_match_optional_group() {
    assert_eq!(
        run_python_one("import re\nm = re.match('colou?r', 'color')\nprint(m.group(0))\n"),
        "color"
    );
}

#[test]
fn re_search_non_greedy_star() {
    assert_eq!(
        run_python_one("import re\nm = re.search('<.*?>', '<a><b>')\nprint(m.group(0))\n"),
        "<a>"
    );
}

#[test]
fn re_findall_digits_with_word_context() {
    assert_eq!(
        run_print("import re; re.findall(r'\\d+', 'order 42 item 7')"),
        "['42', '7']"
    );
}

#[test]
fn re_split_on_dash() {
    assert_eq!(
        run_print("import re; re.split('-', '2024-06-24')"),
        "['2024', '06', '24']"
    );
}

#[test]
fn re_search_span_tuple() {
    assert_eq!(
        run_python_one("import re\nm = re.search('cd', 'abcd')\nprint(m.span())\n"),
        "(2, 4)"
    );
}

