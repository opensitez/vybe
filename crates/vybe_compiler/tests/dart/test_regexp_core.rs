//! RegExp: hasMatch, firstMatch groups, allMatches count, flags, escapes.

dart_cases! {
    regexp_has_match_true_on_digit_run => {
        r#"void main() {
  var re = RegExp(r'\d+');
  print(re.hasMatch('abc123'));
}"#,
        ["true"]
    };

    regexp_has_match_false_without_pattern => {
        r#"void main() {
  var re = RegExp(r'\d+');
  print(re.hasMatch('letters'));
}"#,
        ["false"]
    };

    regexp_has_match_true_on_exact_word => {
        r#"void main() {
  var re = RegExp(r'^hello$');
  print(re.hasMatch('hello'));
}"#,
        ["true"]
    };

    regexp_has_match_false_when_extra_suffix => {
        r#"void main() {
  var re = RegExp(r'^hello$');
  print(re.hasMatch('hello!'));
}"#,
        ["false"]
    };

    regexp_has_match_empty_pattern_on_empty_string => {
        r#"void main() {
  var re = RegExp('');
  print(re.hasMatch(''));
}"#,
        ["true"]
    };

    regexp_has_match_empty_pattern_on_non_empty => {
        r#"void main() {
  var re = RegExp('');
  print(re.hasMatch('x'));
}"#,
        ["true"]
    };

    regexp_first_match_group_zero_full_hit => {
        r#"void main() {
  var re = RegExp(r'\d+');
  var m = re.firstMatch('x99y');
  print(m!.group(0));
}"#,
        ["99"]
    };

    regexp_first_match_group_one_first_capture => {
        r#"void main() {
  var re = RegExp(r'(\d+)-(\d+)');
  var m = re.firstMatch('12-34');
  print(m!.group(1));
}"#,
        ["12"]
    };

    regexp_first_match_group_two_second_capture => {
        r#"void main() {
  var re = RegExp(r'(\d+)-(\d+)');
  var m = re.firstMatch('12-34');
  print(m!.group(2));
}"#,
        ["34"]
    };

    regexp_first_match_returns_null_when_absent => {
        r#"void main() {
  var re = RegExp(r'\d+');
  print(re.firstMatch('abc') == null);
}"#,
        ["true"]
    };

    regexp_first_match_word_capture => {
        r#"void main() {
  var re = RegExp(r'(\w+)@(\w+)');
  var m = re.firstMatch('user@host');
  print(m!.group(1));
}"#,
        ["user"]
    };

    regexp_first_match_domain_capture => {
        r#"void main() {
  var re = RegExp(r'(\w+)@(\w+)');
  var m = re.firstMatch('user@host');
  print(m!.group(2));
}"#,
        ["host"]
    };

    regexp_first_match_alternation_left_branch => {
        r#"void main() {
  var re = RegExp(r'(cat|dog)');
  var m = re.firstMatch('the dog ran');
  print(m!.group(1));
}"#,
        ["dog"]
    };

    regexp_all_matches_count_digits => {
        r#"void main() {
  var re = RegExp(r'\d');
  var count = 0;
  for (var _ in re.allMatches('a1b22c3')) {
    count++;
  }
  print(count);
}"#,
        ["3"]
    };

    regexp_all_matches_count_word_runs => {
        r#"void main() {
  var re = RegExp(r'\w+');
  var count = 0;
  for (var _ in re.allMatches('one two three')) {
    count++;
  }
  print(count);
}"#,
        ["3"]
    };

    regexp_all_matches_zero_on_no_hits => {
        r#"void main() {
  var re = RegExp(r'\d+');
  var count = 0;
  for (var _ in re.allMatches('nodigits')) {
    count++;
  }
  print(count);
}"#,
        ["0"]
    };

    regexp_all_matches_count_via_to_list_length => {
        r#"void main() {
  var re = RegExp(r'\d+');
  print(re.allMatches('a1b22c333').toList().length);
}"#,
        ["3"]
    };

    regexp_all_matches_group_zero_each_hit => {
        r#"void main() {
  var re = RegExp(r'\d+');
  var sum = 0;
  for (var m in re.allMatches('a1b22')) {
    sum += int.parse(m.group(0)!);
  }
  print(sum);
}"#,
        ["23"]
    };

    regexp_case_sensitive_default_rejects_uppercase => {
        r#"void main() {
  var re = RegExp(r'abc');
  print(re.hasMatch('ABC'));
}"#,
        ["false"]
    };

    regexp_case_sensitive_matches_exact_case => {
        r#"void main() {
  var re = RegExp(r'abc');
  print(re.hasMatch('abc'));
}"#,
        ["true"]
    };

    regexp_case_insensitive_flag_matches_uppercase => {
        r#"void main() {
  var re = RegExp(r'abc', caseSensitive: false);
  print(re.hasMatch('ABC'));
}"#,
        ["true"]
    };

    regexp_case_insensitive_flag_matches_mixed_case => {
        r#"void main() {
  var re = RegExp(r'abc', caseSensitive: false);
  print(re.hasMatch('AbC'));
}"#,
        ["true"]
    };

    regexp_case_insensitive_first_match_group => {
        r#"void main() {
  var re = RegExp(r'(ab)c', caseSensitive: false);
  var m = re.firstMatch('xxAbCyy');
  print(m!.group(1));
}"#,
        ["AbC"]
    };

    regexp_multiline_anchor_matches_second_line => {
        r#"void main() {
  var re = RegExp(r'^world', multiLine: true);
  print(re.hasMatch('hello\nworld'));
}"#,
        ["true"]
    };

    regexp_without_multiline_anchor_misses_second_line => {
        r#"void main() {
  var re = RegExp(r'^world');
  print(re.hasMatch('hello\nworld'));
}"#,
        ["false"]
    };

    regexp_multiline_dollar_anchor_end_of_line => {
        r#"void main() {
  var re = RegExp(r'world$', multiLine: true);
  print(re.hasMatch('hello\nworld'));
}"#,
        ["true"]
    };

    regexp_multiline_counts_line_anchors => {
        r#"void main() {
  var re = RegExp(r'^\d$', multiLine: true);
  var count = 0;
  for (var _ in re.allMatches('1\n2\n3')) {
    count++;
  }
  print(count);
}"#,
        ["3"]
    };

    regexp_escape_dot_matches_literal_period => {
        r#"void main() {
  var re = RegExp(r'\.');
  print(re.hasMatch('a.b'));
}"#,
        ["true"]
    };

    regexp_escape_dot_does_not_match_letter => {
        r#"void main() {
  var re = RegExp(r'\.');
  print(re.hasMatch('axb'));
}"#,
        ["false"]
    };

    regexp_escape_backslash_digit_class => {
        r#"void main() {
  var re = RegExp(r'\d+');
  var m = re.firstMatch('id:42');
  print(m!.group(0));
}"#,
        ["42"]
    };

    regexp_escape_word_boundary => {
        r#"void main() {
  var re = RegExp(r'\bcat\b');
  print(re.hasMatch('the cat sat'));
}"#,
        ["true"]
    };

    regexp_escape_word_boundary_rejects_substring => {
        r#"void main() {
  var re = RegExp(r'\bcat\b');
  print(re.hasMatch('concatenate'));
}"#,
        ["false"]
    };

    regexp_raw_string_backslash_digit => {
        r#"void main() {
  var re = RegExp(r'\d{2}');
  var m = re.firstMatch('z99z');
  print(m!.group(0));
}"#,
        ["99"]
    };

    regexp_normal_string_doubled_backslash_digit => {
        r#"void main() {
  var re = RegExp('\\d+');
  var m = re.firstMatch('n7n');
  print(m!.group(0));
}"#,
        ["7"]
    };

    regexp_character_class_vowels => {
        r#"void main() {
  var re = RegExp(r'[aeiou]');
  var m = re.firstMatch('bcdfg');
  print(m == null);
}"#,
        ["true"]
    };

    regexp_character_class_matches_first_vowel => {
        r#"void main() {
  var re = RegExp(r'[aeiou]');
  var m = re.firstMatch('bcda');
  print(m!.group(0));
}"#,
        ["a"]
    };

    regexp_negated_character_class => {
        r#"void main() {
  var re = RegExp(r'[^0-9]+');
  var m = re.firstMatch('123abc456');
  print(m!.group(0));
}"#,
        ["abc"]
    };

    regexp_quantifier_exact_three_digits => {
        r#"void main() {
  var re = RegExp(r'\d{3}');
  var m = re.firstMatch('a1234');
  print(m!.group(0));
}"#,
        ["123"]
    };

    regexp_quantifier_one_or_more_letters => {
        r#"void main() {
  var re = RegExp(r'[a-z]+');
  var m = re.firstMatch('1hello2');
  print(m!.group(0));
}"#,
        ["hello"]
    };

    regexp_optional_question_mark => {
        r#"void main() {
  var re = RegExp(r'colou?r');
  print(re.hasMatch('color'));
}"#,
        ["true"]
    };

    regexp_optional_question_mark_u_form => {
        r#"void main() {
  var re = RegExp(r'colou?r');
  print(re.hasMatch('colour'));
}"#,
        ["true"]
    };

    regexp_alternation_pipe => {
        r#"void main() {
  var re = RegExp(r'cat|dog');
  print(re.hasMatch('my dog'));
}"#,
        ["true"]
    };

    regexp_group_non_capturing_quantifier => {
        r#"void main() {
  var re = RegExp(r'(?:ab)+');
  var m = re.firstMatch('xxababab');
  print(m!.group(0));
}"#,
        ["ababab"]
    };

    regexp_has_match_after_string_starts_with => {
        r#"void main() {
  var re = RegExp(r'^start');
  print(re.hasMatch('start here'));
}"#,
        ["true"]
    };

    regexp_has_match_before_string_ends_with => {
        r#"void main() {
  var re = RegExp(r'end$');
  print(re.hasMatch('the end'));
}"#,
        ["true"]
    };

    regexp_first_match_at_start_index => {
        r#"void main() {
  var re = RegExp(r'\d+');
  var m = re.firstMatch('42abc');
  print(m!.start);
}"#,
        ["0"]
    };

    regexp_all_matches_on_repeated_pattern => {
        r#"void main() {
  var re = RegExp(r'ab');
  print(re.allMatches('ababab').toList().length);
}"#,
        ["3"]
    };

    regexp_case_insensitive_all_matches_count => {
        r#"void main() {
  var re = RegExp(r'a', caseSensitive: false);
  print(re.allMatches('AaA').toList().length);
}"#,
        ["3"]
    };

    regexp_multiline_first_match_on_second_line => {
        r#"void main() {
  var re = RegExp(r'^line', multiLine: true);
  var m = re.firstMatch('first\nline2');
  print(m!.group(0));
}"#,
        ["line"]
    };

    regexp_escape_parentheses_literal => {
        r#"void main() {
  var re = RegExp(r'\(x\)');
  print(re.hasMatch('(x)'));
}"#,
        ["true"]
    };

    regexp_escape_bracket_literal => {
        r#"void main() {
  var re = RegExp(r'\[item\]');
  print(re.hasMatch('[item]'));
}"#,
        ["true"]
    };

    regexp_named_capture_style_group_numbering => {
        r#"void main() {
  var re = RegExp(r'(a)(b)(c)');
  var m = re.firstMatch('xyzabc');
  print(m!.group(3));
}"#,
        ["c"]
    };

    regexp_has_match_unicode_letter_class => {
        r#"void main() {
  var re = RegExp(r'\p{L}+', unicode: true);
  print(re.hasMatch('café'));
}"#,
        ["true"]
    };

    regexp_all_matches_skips_non_overlapping => {
        r#"void main() {
  var re = RegExp(r'aa');
  print(re.allMatches('aaaa').toList().length);
}"#,
        ["2"]
    };

    regexp_first_match_on_only_digits_string => {
        r#"void main() {
  var re = RegExp(r'^\d+$');
  print(re.hasMatch('12345'));
}"#,
        ["true"]
    };

    regexp_has_match_whitespace_class => {
        r#"void main() {
  var re = RegExp(r'\s+');
  print(re.hasMatch('a  b'));
}"#,
        ["true"]
    };

    regexp_first_match_hex_class => {
        r#"void main() {
  var re = RegExp(r'[0-9a-f]+');
  var m = re.firstMatch('zzff00');
  print(m!.group(0));
}"#,
        ["ff"]
    };
}
