use crate::helpers::{run_print, run_python_one};

#[test]
fn partition_splits_on_first_separator() {
    assert_eq!(run_print("'a.b.c'.partition('.')"), "('a', '.', 'b.c')");
}

#[test]
fn rpartition_splits_on_last_separator() {
    assert_eq!(run_print("'a.b.c'.rpartition('.')"), "('a.b', '.', 'c')");
}

#[test]
fn partition_missing_sep_returns_head_empty_tail() {
    assert_eq!(run_print("'abc'.partition('-')"), "('abc', '', '')");
}

#[test]
fn rpartition_missing_sep_returns_empty_head_tail() {
    assert_eq!(run_print("'abc'.rpartition('-')"), "('', '', 'abc')");
}

#[test]
fn partition_sep_at_start() {
    assert_eq!(run_print("'.abc'.partition('.')"), "('', '.', 'abc')");
}

#[test]
fn rpartition_sep_at_end() {
    assert_eq!(run_print("'abc.'.rpartition('.')"), "('abc', '.', '')");
}

#[test]
fn split_default_whitespace() {
    assert_eq!(run_print("'a  b\\tc'.split()"), "['a', 'b', 'c']");
}

#[test]
fn split_explicit_space_single() {
    assert_eq!(run_print("'a b c'.split(' ')"), "['a', 'b', 'c']");
}

#[test]
fn split_maxsplit_one() {
    assert_eq!(run_print("'a-b-c'.split('-', 1)"), "['a', 'b-c']");
}

#[test]
fn rsplit_maxsplit_one() {
    assert_eq!(run_print("'a-b-c'.rsplit('-', 1)"), "['a-b', 'c']");
}

#[test]
fn splitlines_unix() {
    assert_eq!(run_print("'a\\nb\\n'.splitlines()"), "['a', 'b']");
}

#[test]
fn splitlines_keepends() {
    assert_eq!(
        run_python_one("print(repr('a\\n'.splitlines(True)))\n"),
        "['a\\n']"
    );
}

#[test]
fn splitlines_no_trailing_empty_without_final_newline() {
    assert_eq!(run_print("'a\\nb'.splitlines()"), "['a', 'b']");
}

#[test]
fn split_on_colon_for_path_like() {
    assert_eq!(
        run_print("'usr:local:bin'.split(':')"),
        "['usr', 'local', 'bin']"
    );
}

#[test]
fn rsplit_on_colon_once() {
    assert_eq!(
        run_print("'usr:local:bin'.rsplit(':', 1)"),
        "['usr:local', 'bin']"
    );
}

#[test]
fn partition_on_multichar_not_special() {
    assert_eq!(run_print("'ab::cd'.partition('::')"), "('ab', '::', 'cd')");
}

#[test]
fn split_empty_string_returns_empty_list() {
    assert_eq!(run_print("''.split()"), "[]");
}

#[test]
fn split_single_char_string_no_delim() {
    assert_eq!(run_print("'x'.split(',')"), "['x']");
}

#[test]
fn partition_single_char_string_no_delim() {
    assert_eq!(run_print("'x'.partition(',')"), "('x', '', '')");
}

#[test]
fn splitlines_empty_string() {
    assert_eq!(run_print("''.splitlines()"), "[]");
}

#[test]
fn split_only_separators_returns_empties() {
    assert_eq!(run_print("','.split(',')"), "['', '']");
}

#[test]
fn rpartition_only_separators() {
    assert_eq!(run_print("'.'.rpartition('.')"), "('', '.', '')");
}

#[test]
fn split_with_tab_delimiter() {
    assert_eq!(run_print("'a\\tb\\tc'.split('\\t')"), "['a', 'b', 'c']");
}

#[test]
fn splitlines_mixed_cr_nl() {
    assert_eq!(run_print("'a\\r\\nb'.splitlines()"), "['a', 'b']");
}

#[test]
fn partition_used_to_strip_prefix() {
    assert_eq!(
        run_python_one("head, sep, tail = 'key=value'.partition('=')\nprint(tail)\n"),
        "value"
    );
}

#[test]
fn rpartition_used_to_strip_suffix() {
    assert_eq!(
        run_python_one("head, sep, tail = 'file.txt.bak'.rpartition('.')\nprint(head)\n"),
        "file.txt"
    );
}

#[test]
fn split_maxsplit_zero_returns_whole() {
    assert_eq!(run_print("'a-b-c'.split('-', 0)"), "['a-b-c']");
}

#[test]
fn rsplit_maxsplit_zero_returns_whole() {
    assert_eq!(run_print("'a-b-c'.rsplit('-', 0)"), "['a-b-c']");
}

#[test]
fn split_on_newline_char() {
    assert_eq!(run_print("'x\\ny'.split('\\n')"), "['x', 'y']");
}

#[test]
fn splitlines_single_line_no_break() {
    assert_eq!(run_print("'hello'.splitlines()"), "['hello']");
}

#[test]
fn partition_unicode_separator() {
    assert_eq!(run_print("'a→b→c'.partition('→')"), "('a', '→', 'b→c')");
}

#[test]
fn split_consecutive_spaces_default_split_collapses() {
    assert_eq!(run_print("'a   b'.split()"), "['a', 'b']");
}

#[test]
fn split_explicit_space_preserves_empty_tokens() {
    assert_eq!(run_print("'a  b'.split(' ')"), "['a', '', 'b']");
}

#[test]
fn rpartition_multichar_last_occurrence() {
    assert_eq!(
        run_print("'foo::bar::baz'.rpartition('::')"),
        "('foo::bar', '::', 'baz')"
    );
}

#[test]
fn splitlines_with_blank_line_in_middle() {
    assert_eq!(run_print("'a\\n\\nb'.splitlines()"), "['a', '', 'b']");
}
