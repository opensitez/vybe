//! Dart `StringBuffer`: write, writeAll, writeln, clear, length, char codes, cascades.

dart_cases! {
    string_buffer_write_single_string => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('hello');
  print(buf.toString());
}"#,
        ["hello"]
    };

    string_buffer_write_concatenates_multiple_calls => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('foo');
  buf.write('bar');
  print(buf.toString());
}"#,
        ["foobar"]
    };

    string_buffer_write_empty_string_no_change => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('x');
  buf.write('');
  print(buf.toString());
  print(buf.length);
}"#,
        ["x", "1"]
    };

    string_buffer_write_all_joins_iterable_without_separator => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeAll(['a', 'b', 'c']);
  print(buf.toString());
}"#,
        ["abc"]
    };

    string_buffer_write_all_with_comma_separator => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeAll(['one', 'two', 'three'], ',');
  print(buf.toString());
}"#,
        ["one,two,three"]
    };

    string_buffer_write_all_with_dash_separator => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeAll(['2024', '06', '26'], '-');
  print(buf.toString());
}"#,
        ["2024-06-26"]
    };

    string_buffer_write_all_empty_iterable => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeAll(<String>[]);
  print(buf.toString());
  print(buf.isEmpty);
}"#,
        ["", "true"]
    };

    string_buffer_write_all_single_element => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeAll(['solo']);
  print(buf.toString());
}"#,
        ["solo"]
    };

    string_buffer_writeln_appends_newline => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeln('line');
  print(buf.toString());
}"#,
        ["line\n"]
    };

    string_buffer_writeln_without_argument_adds_blank_line => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeln();
  print(buf.toString());
  print(buf.length);
}"#,
        ["\n", "1"]
    };

    string_buffer_write_then_writeln => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('prefix');
  buf.writeln('suffix');
  print(buf.toString());
}"#,
        ["prefixsuffix\n"]
    };

    string_buffer_multiple_writeln_lines => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeln('a');
  buf.writeln('b');
  print(buf.toString());
}"#,
        ["a\nb\n"]
    };

    string_buffer_clear_empties_content => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('remove me');
  buf.clear();
  print(buf.toString());
  print(buf.length);
  print(buf.isEmpty);
}"#,
        ["", "0", "true"]
    };

    string_buffer_clear_then_write_restarts_content => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('old');
  buf.clear();
  buf.write('new');
  print(buf.toString());
}"#,
        ["new"]
    };

    string_buffer_length_tracks_characters => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('abcd');
  print(buf.length);
}"#,
        ["4"]
    };

    string_buffer_length_after_writeln_includes_newline => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeln('hi');
  print(buf.length);
}"#,
        ["3"]
    };

    string_buffer_is_empty_on_creation => {
        r#"void main() {
  var buf = StringBuffer();
  print(buf.isEmpty);
  print(!buf.isEmpty);
}"#,
        ["true", "false"]
    };

    string_buffer_is_not_empty_after_write => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('x');
  print(buf.isEmpty);
  print(!buf.isEmpty);
}"#,
        ["false", "true"]
    };

    string_buffer_to_string_matches_written_content => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('dart');
  buf.write('lang');
  print(buf.toString());
}"#,
        ["dartlang"]
    };

    string_buffer_write_char_code_lowercase_a => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeCharCode(97);
  print(buf.toString());
}"#,
        ["a"]
    };

    string_buffer_write_char_code_uppercase_z => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeCharCode(90);
  print(buf.toString());
}"#,
        ["Z"]
    };

    string_buffer_write_char_code_digit_zero => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeCharCode(48);
  print(buf.toString());
}"#,
        ["0"]
    };

    string_buffer_write_char_codes_build_word => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeCharCode(104);
  buf.writeCharCode(105);
  print(buf.toString());
}"#,
        ["hi"]
    };

    string_buffer_write_char_code_after_string_write => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('A');
  buf.writeCharCode(66);
  print(buf.toString());
}"#,
        ["AB"]
    };

    string_buffer_cascade_write_chain => {
        r#"void main() {
  var buf = StringBuffer()
    ..write('x')
    ..write('y')
    ..write('z');
  print(buf.toString());
}"#,
        ["xyz"]
    };

    string_buffer_cascade_write_all_then_write => {
        r#"void main() {
  var buf = StringBuffer()
    ..writeAll(['a', 'b'], '-')
    ..write('!');
  print(buf.toString());
}"#,
        ["a-b!"]
    };

    string_buffer_cascade_writeln_chain => {
        r#"void main() {
  var buf = StringBuffer()
    ..writeln('first')
    ..writeln('second');
  print(buf.toString());
}"#,
        ["first\nsecond\n"]
    };

    string_buffer_cascade_clear_then_write => {
        r#"void main() {
  var buf = StringBuffer()
    ..write('discard')
    ..clear()
    ..write('kept');
  print(buf.toString());
}"#,
        ["kept"]
    };

    string_buffer_cascade_returns_same_receiver => {
        r#"void main() {
  var buf = StringBuffer();
  var same = buf..write('a')..write('b');
  print(same == buf);
  print(buf.toString());
}"#,
        ["true", "ab"]
    };

    string_buffer_cascade_write_char_code_chain => {
        r#"void main() {
  var buf = StringBuffer()
    ..writeCharCode(68)
    ..writeCharCode(69)
    ..write('F');
  print(buf.toString());
}"#,
        ["DEF"]
    };

    string_buffer_write_int_converts_to_string => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write(42);
  print(buf.toString());
}"#,
        ["42"]
    };

    string_buffer_write_bool_converts_to_string => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write(true);
  print(buf.toString());
}"#,
        ["true"]
    };

    string_buffer_write_all_int_list => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeAll([1, 2, 3], '');
  print(buf.toString());
}"#,
        ["123"]
    };

    string_buffer_mixed_write_and_write_all => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('start-');
  buf.writeAll(['mid', 'end'], ':');
  print(buf.toString());
}"#,
        ["start-mid:end"]
    };

    string_buffer_length_zero_after_clear => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('long content here');
  buf.clear();
  print(buf.length);
  print(buf.isEmpty);
}"#,
        ["0", "true"]
    };

    string_buffer_to_string_after_clear_is_empty => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('gone');
  buf.clear();
  print(buf.toString() == '');
}"#,
        ["true"]
    };

    string_buffer_write_unicode_character => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('\u2665');
  print(buf.toString());
  print(buf.length);
}"#,
        ["\u{2665}", "1"]
    };

    string_buffer_write_all_preserves_order => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeAll(['first', 'second', 'third']);
  print(buf.toString());
}"#,
        ["firstsecondthird"]
    };

    string_buffer_writeln_then_write_no_extra_newline_between => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeln('line1');
  buf.write('line2');
  print(buf.toString());
}"#,
        ["line1\nline2"]
    };

    string_buffer_repeated_clear_is_idempotent => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('data');
  buf.clear();
  buf.clear();
  print(buf.length);
  print(buf.toString());
}"#,
        ["0", ""]
    };

    string_buffer_cascade_mixed_write_writeln_write_char_code => {
        r#"void main() {
  var buf = StringBuffer()
    ..write('A')
    ..writeln('B')
    ..writeCharCode(67);
  print(buf.toString());
}"#,
        ["AB\nC"]
    };

    string_buffer_write_all_with_empty_separator => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeAll(['x', 'y', 'z'], '');
  print(buf.toString());
}"#,
        ["xyz"]
    };

    string_buffer_write_char_code_space => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeCharCode(32);
  buf.writeCharCode(32);
  print(buf.toString().length);
  print(buf.length);
}"#,
        ["2", "2"]
    };

    string_buffer_length_grows_with_each_write => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('a');
  print(buf.length);
  buf.write('bc');
  print(buf.length);
}"#,
        ["1", "3"]
    };

    string_buffer_write_all_after_clear => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('old');
  buf.clear();
  buf.writeAll(['fresh'], '');
  print(buf.toString());
}"#,
        ["fresh"]
    };

    string_buffer_cascade_length_after_writes => {
        r#"void main() {
  var buf = StringBuffer()
    ..write('12345')
    ..write('67890');
  print(buf.length);
  print(buf.toString().length);
}"#,
        ["10", "10"]
    };

    string_buffer_write_null_string_literal_empty => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('');
  buf.write('ok');
  print(buf.toString());
}"#,
        ["ok"]
    };

    string_buffer_write_all_single_char_elements => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeAll(['x', 'y'], '|');
  print(buf.toString());
}"#,
        ["x|y"]
    };

    string_buffer_writeln_multiple_empty_lines => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeln();
  buf.writeln();
  print(buf.length);
  print(buf.toString());
}"#,
        ["2", "\n\n"]
    };

    string_buffer_write_char_code_before_clear_removed => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeCharCode(88);
  buf.clear();
  buf.write('Y');
  print(buf.toString());
}"#,
        ["Y"]
    };

    string_buffer_cascade_write_all_writeln_then_read => {
        r#"void main() {
  var buf = StringBuffer()
    ..writeAll(['a', 'b'], '')
    ..writeln('c');
  print(buf.toString());
  print(buf.length);
}"#,
        ["abc\n", "4"]
    };

    string_buffer_write_special_characters => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('tab\there');
  print(buf.toString().contains('\t'));
}"#,
        ["true"]
    };

    string_buffer_is_not_empty_after_write_all => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeAll(['only']);
  print(!buf.isEmpty);
  print(buf.isEmpty);
}"#,
        ["true", "false"]
    };

    string_buffer_write_then_toString_does_not_consume => {
        r#"void main() {
  var buf = StringBuffer();
  buf.write('persist');
  print(buf.toString());
  print(buf.toString());
  print(buf.length);
}"#,
        ["persist", "persist", "7"]
    };

    string_buffer_write_char_code_newline => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeCharCode(10);
  print(buf.length);
  print(buf.toString() == '\n');
}"#,
        ["1", "true"]
    };

    string_buffer_write_all_then_clear_then_writeln => {
        r#"void main() {
  var buf = StringBuffer();
  buf.writeAll(['old', 'data']);
  buf.clear();
  buf.writeln('fresh');
  print(buf.toString());
  print(buf.isEmpty);
}"#,
        ["fresh\n", "false"]
    };
}
