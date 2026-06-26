//! Core String instance methods: search, slice, transform, pad, compare, runes.

dart_cases! {
    substring_with_start_and_end => {
        r#"void main() {
  print('hello'.substring(1, 4));
}"#,
        ["ell"]
    };

    substring_start_only_to_end => {
        r#"void main() {
  print('hello'.substring(2));
}"#,
        ["llo"]
    };

    substring_zero_to_full_length => {
        r#"void main() {
  print('dart'.substring(0, 4));
}"#,
        ["dart"]
    };

    substring_empty_range_at_start => {
        r#"void main() {
  print('abc'.substring(0, 0).isEmpty);
}"#,
        ["true"]
    };

    indexOf_finds_substring_at_start => {
        r#"void main() {
  print('hello'.indexOf('he'));
}"#,
        ["0"]
    };

    indexOf_finds_substring_in_middle => {
        r#"void main() {
  print('hello world'.indexOf('world'));
}"#,
        ["6"]
    };

    indexOf_returns_negative_one_when_missing => {
        r#"void main() {
  print('hello'.indexOf('xyz'));
}"#,
        ["-1"]
    };

    indexOf_empty_needle_at_start => {
        r#"void main() {
  print('abc'.indexOf(''));
}"#,
        ["0"]
    };

    indexOf_with_start_position => {
        r#"void main() {
  print('banana'.indexOf('na', 2));
}"#,
        ["3"]
    };

    startsWith_true_for_matching_prefix => {
        r#"void main() {
  print('hello'.startsWith('he'));
}"#,
        ["true"]
    };

    startsWith_false_for_wrong_prefix => {
        r#"void main() {
  print('hello'.startsWith('lo'));
}"#,
        ["false"]
    };

    startsWith_with_position_offset => {
        r#"void main() {
  print('hello'.startsWith('ll', 2));
}"#,
        ["true"]
    };

    endsWith_true_for_matching_suffix => {
        r#"void main() {
  print('hello.dart'.endsWith('.dart'));
}"#,
        ["true"]
    };

    endsWith_false_for_wrong_suffix => {
        r#"void main() {
  print('hello'.endsWith('he'));
}"#,
        ["false"]
    };

    replaceAll_replaces_every_occurrence => {
        r#"void main() {
  print('aaa'.replaceAll('a', 'b'));
}"#,
        ["bbb"]
    };

    replaceAll_single_character_substitution => {
        r#"void main() {
  print('hello'.replaceAll('l', 'r'));
}"#,
        ["herro"]
    };

    replaceAll_no_match_leaves_string_unchanged => {
        r#"void main() {
  print('hello'.replaceAll('z', 'x'));
}"#,
        ["hello"]
    };

    replaceAll_empty_replacement_deletes_matches => {
        r#"void main() {
  print('a-b-c'.replaceAll('-', ''));
}"#,
        ["abc"]
    };

    split_comma_delimited_returns_three_parts => {
        r#"void main() {
  var parts = 'a,b,c'.split(',');
  print(parts.length);
}"#,
        ["3"]
    };

    split_single_element_no_delimiter => {
        r#"void main() {
  var parts = 'solo'.split(',');
  print(parts.length);
}"#,
        ["1"]
    };

    split_empty_string_yields_one_empty_part => {
        r#"void main() {
  var parts = ''.split(',');
  print(parts.length);
}"#,
        ["1"]
    };

    split_space_delimited_words => {
        r#"void main() {
  var parts = 'one two three'.split(' ');
  print(parts.join('-'));
}"#,
        ["one-two-three"]
    };

    trim_removes_leading_and_trailing_whitespace => {
        r#"void main() {
  print('  hi  '.trim());
}"#,
        ["hi"]
    };

    trimLeft_removes_leading_whitespace_only => {
        r#"void main() {
  print('  hi  '.trimLeft());
}"#,
        ["hi  "]
    };

    trimRight_removes_trailing_whitespace_only => {
        r#"void main() {
  print('  hi  '.trimRight());
}"#,
        ["  hi"]
    };

    padLeft_pads_with_zeros_to_width => {
        r#"void main() {
  print('7'.padLeft(3, '0'));
}"#,
        ["007"]
    };

    padLeft_already_at_width_unchanged => {
        r#"void main() {
  print('abc'.padLeft(3, '0'));
}"#,
        ["abc"]
    };

    padLeft_longer_than_width_unchanged => {
        r#"void main() {
  print('abcd'.padLeft(3, '0'));
}"#,
        ["abcd"]
    };

    padRight_pads_with_dots_to_width => {
        r#"void main() {
  print('hi'.padRight(5, '.'));
}"#,
        ["hi..."]
    };

    padRight_default_space_padding => {
        r#"void main() {
  print('x'.padRight(3));
}"#,
        ["x  "]
    };

    toUpperCase_converts_lowercase_letters => {
        r#"void main() {
  print('hello'.toUpperCase());
}"#,
        ["HELLO"]
    };

    toUpperCase_leaves_digits_unchanged => {
        r#"void main() {
  print('ab12'.toUpperCase());
}"#,
        ["AB12"]
    };

    toLowerCase_converts_uppercase_letters => {
        r#"void main() {
  print('HELLO'.toLowerCase());
}"#,
        ["hello"]
    };

    toLowerCase_on_mixed_case_string => {
        r#"void main() {
  print('HeLLo'.toLowerCase());
}"#,
        ["hello"]
    };

    contains_true_for_present_substring => {
        r#"void main() {
  print('hello'.contains('ell'));
}"#,
        ["true"]
    };

    contains_false_for_absent_substring => {
        r#"void main() {
  print('hello'.contains('xyz'));
}"#,
        ["false"]
    };

    contains_empty_string_is_always_true => {
        r#"void main() {
  print('hello'.contains(''));
}"#,
        ["true"]
    };

    compareTo_equal_strings_returns_zero => {
        r#"void main() {
  print('abc'.compareTo('abc'));
}"#,
        ["0"]
    };

    compareTo_less_than_returns_negative => {
        r#"void main() {
  print('apple'.compareTo('banana'));
}"#,
        ["-1"]
    };

    compareTo_greater_than_returns_positive => {
        r#"void main() {
  print('zebra'.compareTo('apple'));
}"#,
        ["1"]
    };

    isEmpty_true_for_empty_literal => {
        r#"void main() {
  print(''.isEmpty);
}"#,
        ["true"]
    };

    isEmpty_false_for_non_empty_string => {
        r#"void main() {
  print('x'.isEmpty);
}"#,
        ["false"]
    };

    isNotEmpty_true_for_non_empty_string => {
        r#"void main() {
  print('x'.isNotEmpty);
}"#,
        ["true"]
    };

    isNotEmpty_false_for_empty_string => {
        r#"void main() {
  print(''.isNotEmpty);
}"#,
        ["false"]
    };

    length_counts_ascii_code_units => {
        r#"void main() {
  print('hello'.length);
}"#,
        ["5"]
    };

    runes_length_matches_length_for_ascii => {
        r#"void main() {
  var s = 'hello';
  print(s.runes.length == s.length);
}"#,
        ["true"]
    };

    runes_length_counts_emoji_as_one => {
        r#"void main() {
  var s = '🙂';
  print(s.runes.length);
}"#,
        ["1"]
    };

    length_counts_emoji_as_two_code_units => {
        r#"void main() {
  var s = '🙂';
  print(s.length);
}"#,
        ["2"]
    };

    runes_length_counts_accented_letter => {
        r#"void main() {
  var s = 'café';
  print(s.runes.length);
}"#,
        ["4"]
    };

    substring_works_on_unicode_string => {
        r#"void main() {
  var s = 'café';
  print(s.substring(0, 3));
}"#,
        ["caf"]
    };

    replaceAll_preserves_unmatched_regions => {
        r#"void main() {
  print('axbxc'.replaceAll('x', 'y'));
}"#,
        ["aybyc"]
    };
}
