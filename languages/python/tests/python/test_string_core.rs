use crate::helpers::{run_print, run_python, run_python_one};

#[test]
fn str_upper_lowercase_letters() {
    assert_eq!(run_print("'hello'.upper()"), "HELLO");
}

#[test]
fn str_lower_uppercase_letters() {
    assert_eq!(run_print("'WORLD'.lower()"), "world");
}

#[test]
fn str_strip_surrounding_whitespace() {
    assert_eq!(run_print("'  trim  '.strip()"), "trim");
}

#[test]
fn str_strip_custom_character_set() {
    assert_eq!(run_print("'***hi***'.strip('*')"), "hi");
}

#[test]
fn str_split_on_comma_separator() {
    assert_eq!(
        run_python("for p in 'a,b,c'.split(','):\n    print(p)\n"),
        vec!["a", "b", "c"]
    );
}

#[test]
fn str_split_maxsplit_limits_parts() {
    assert_eq!(
        run_python("for p in 'a:b:c:d'.split(':', 2):\n    print(p)\n"),
        vec!["a", "b", "c:d"]
    );
}

#[test]
fn str_split_whitespace_default() {
    assert_eq!(
        run_python("for p in 'one two  three'.split():\n    print(p)\n"),
        vec!["one", "two", "three"]
    );
}

#[test]
fn str_join_with_dash_separator() {
    assert_eq!(run_print("'-'.join(['x', 'y', 'z'])"), "x-y-z");
}

#[test]
fn str_replace_all_matches() {
    assert_eq!(run_print("'banana'.replace('a', 'o')"), "bonono");
}

#[test]
fn str_replace_count_limits_replacements() {
    assert_eq!(run_print("'banana'.replace('a', 'o', 1)"), "bonana");
}

#[test]
fn str_find_existing_substring() {
    assert_eq!(run_print("'abcdef'.find('cd')"), "2");
}

#[test]
fn str_find_missing_substring_returns_negative_one() {
    assert_eq!(run_print("'abcdef'.find('z')"), "-1");
}

#[test]
fn str_index_existing_substring() {
    assert_eq!(run_print("'abcdef'.index('cd')"), "2");
}

#[test]
fn str_count_overlapping_not_counted() {
    assert_eq!(run_print("'aaaa'.count('aa')"), "2");
}

#[test]
fn str_startswith_matching_prefix() {
    assert_eq!(run_print("'python'.startswith('py')"), "True");
}

#[test]
fn str_startswith_non_matching_prefix() {
    assert_eq!(run_print("'python'.startswith('ja')"), "False");
}

#[test]
fn str_startswith_with_start_offset() {
    assert_eq!(run_print("'banana'.startswith('na', 2)"), "True");
}

#[test]
fn str_endswith_matching_suffix() {
    assert_eq!(run_print("'filename.txt'.endswith('.txt')"), "True");
}

#[test]
fn str_endswith_non_matching_suffix() {
    assert_eq!(run_print("'filename.txt'.endswith('.py')"), "False");
}

#[test]
fn str_endswith_with_end_bound() {
    assert_eq!(run_print("'banana'.endswith('an', 0, 3)"), "True");
}

#[test]
fn str_isalpha_all_letters() {
    assert_eq!(run_print("'Alpha'.isalpha()"), "True");
}

#[test]
fn str_isalpha_with_digit_returns_false() {
    assert_eq!(run_print("'a1'.isalpha()"), "False");
}

#[test]
fn str_isdigit_all_decimal_digits() {
    assert_eq!(run_print("'9042'.isdigit()"), "True");
}

#[test]
fn str_isdigit_with_letter_returns_false() {
    assert_eq!(run_print("'9a'.isdigit()"), "False");
}

#[test]
fn str_format_named_keyword_arguments() {
    assert_eq!(
        run_python_one("print('{name}={value}'.format(name='port', value=80))\n"),
        "port=80"
    );
}

#[test]
fn str_format_positional_indices() {
    assert_eq!(
        run_python_one("print('{0}-{1}-{0}'.format('a', 'b'))\n"),
        "a-b-a"
    );
}

#[test]
fn str_format_literal_braces_escaped() {
    assert_eq!(run_python_one("print('{{ok}}'.format())\n"), "{ok}");
}

#[test]
fn str_zfill_pads_with_leading_zeros() {
    assert_eq!(run_print("'42'.zfill(5)"), "00042");
}

#[test]
fn str_zfill_longer_than_input_unchanged() {
    assert_eq!(run_print("'hello'.zfill(3)"), "hello");
}

#[test]
fn str_center_pads_both_sides() {
    assert_eq!(run_print("'hi'.center(6)"), "  hi  ");
}

#[test]
fn str_center_with_custom_fill_character() {
    assert_eq!(run_print("'hi'.center(6, '-')"), "--hi--");
}

#[test]
fn str_ljust_left_aligns_with_padding() {
    assert_eq!(run_print("'go'.ljust(5, '.')"), "go...");
}

#[test]
fn str_rjust_right_aligns_with_padding() {
    assert_eq!(run_print("'go'.rjust(5, '.')"), "...go");
}

#[test]
fn str_partition_splits_on_first_separator() {
    assert_eq!(
        run_python_one("print('key=value'.partition('=')[1])\n"),
        "="
    );
}

#[test]
fn str_partition_before_part_when_found() {
    assert_eq!(
        run_python_one("print('key=value'.partition('=')[0])\n"),
        "key"
    );
}

#[test]
fn str_partition_after_part_when_missing() {
    assert_eq!(run_python_one("print('nosep'.partition('=')[2])\n"), "");
}

#[test]
fn str_rpartition_splits_on_last_separator() {
    assert_eq!(run_python_one("print('a/b/c'.rpartition('/')[2])\n"), "c");
}

#[test]
fn str_rpartition_before_part_from_right() {
    assert_eq!(run_python_one("print('a/b/c'.rpartition('/')[0])\n"), "a/b");
}

#[test]
fn str_removeprefix_strips_matching_prefix() {
    assert_eq!(run_print("'HelloWorld'.removeprefix('Hello')"), "World");
}

#[test]
fn str_removeprefix_no_match_returns_original() {
    assert_eq!(run_print("'HelloWorld'.removeprefix('Bye')"), "HelloWorld");
}

#[test]
fn str_removesuffix_strips_matching_suffix() {
    assert_eq!(
        run_print("'archive.tar.gz'.removesuffix('.gz')"),
        "archive.tar"
    );
}

#[test]
fn str_removesuffix_no_match_returns_original() {
    assert_eq!(
        run_print("'archive.tar.gz'.removesuffix('.zip')"),
        "archive.tar.gz"
    );
}

#[test]
fn str_casefold_uppercase_to_lowercase() {
    assert_eq!(run_print("'Straße'.casefold()"), "straße");
}

#[test]
fn str_expandtabs_replaces_tab_default_width() {
    assert_eq!(run_print("'a\\tb'.expandtabs()"), "a       b");
}

#[test]
fn str_expandtabs_custom_tab_stop() {
    assert_eq!(run_print("'a\\tb'.expandtabs(4)"), "a   b");
}

#[test]
fn str_encode_ascii_roundtrip_via_decode() {
    assert_eq!(
        run_python_one("print('vybe'.encode('ascii').decode('ascii'))\n"),
        "vybe"
    );
}

#[test]
fn str_decode_ascii_from_byte_literal() {
    assert_eq!(run_python_one("print(b'ascii'.decode('ascii'))\n"), "ascii");
}

#[test]
fn str_title_capitalizes_each_word() {
    assert_eq!(run_print("'hello world'.title()"), "Hello World");
}

#[test]
fn str_capitalize_first_char_only() {
    assert_eq!(run_print("'hELLO'.capitalize()"), "Hello");
}

#[test]
fn str_swapcase_inverts_letter_case() {
    assert_eq!(run_print("'PyThOn'.swapcase()"), "pYtHoN");
}
