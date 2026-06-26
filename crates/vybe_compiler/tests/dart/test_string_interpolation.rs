//! String literals, interpolation, adjacency, raw/triple-quoted strings, escapes.

dart_cases! {
    simple_dollar_var_interpolation => {
        r#"void main() {
  var name = 'Ada';
  print('Hello $name');
}"#,
        ["Hello Ada"]
    };

    dollar_var_with_trailing_text => {
        r#"void main() {
  var n = 7;
  print('value=$n!');
}"#,
        ["value=7!"]
    };

    dollar_var_with_leading_text => {
        r#"void main() {
  var x = 2;
  print('${x}items');
}"#,
        ["2items"]
    };

    expr_interpolation_addition => {
        r#"void main() {
  var a = 10;
  var b = 32;
  print('sum=${a + b}');
}"#,
        ["sum=42"]
    };

    expr_interpolation_multiplication => {
        r#"void main() {
  var n = 6;
  print('square=${n * n}');
}"#,
        ["square=36"]
    };

    expr_interpolation_integer_division => {
        r#"void main() {
  var n = 17;
  print('half=${n ~/ 2}');
}"#,
        ["half=8"]
    };

    expr_interpolation_modulo => {
        r#"void main() {
  var n = 17;
  print('mod=${n % 5}');
}"#,
        ["mod=2"]
    };

    expr_interpolation_ternary => {
        r#"void main() {
  var n = 4;
  print('parity=${n % 2 == 0 ? 'even' : 'odd'}');
}"#,
        ["parity=even"]
    };

    expr_interpolation_method_call => {
        r#"void main() {
  var word = 'dart';
  print('upper=${word.toUpperCase()}');
}"#,
        ["upper=DART"]
    };

    multiple_dollar_vars_one_string => {
        r#"void main() {
  var first = 'A';
  var last = 'B';
  print('$first$last');
}"#,
        ["AB"]
    };

    mixed_var_and_expr_interpolation => {
        r#"void main() {
  var x = 3;
  print('x=$x next=${x + 1}');
}"#,
        ["x=3 next=4"]
    };

    adjacent_single_quoted_strings_concat => {
        r#"void main() {
  print('hel' 'lo');
}"#,
        ["hello"]
    };

    adjacent_double_quoted_strings_concat => {
        r#"void main() {
  print("good" "bye");
}"#,
        ["goodbye"]
    };

    adjacent_mixed_quote_strings_concat => {
        r#"void main() {
  print('mix' "ed");
}"#,
        ["mixed"]
    };

    three_adjacent_string_literals_concat => {
        r#"void main() {
  print('a' 'b' 'c');
}"#,
        ["abc"]
    };

    triple_single_quoted_multiline_has_three_lines => {
        r#"void main() {
  var text = '''one
two
three''';
  print(text.split('\n').length);
}"#,
        ["3"]
    };

    triple_double_quoted_multiline_has_two_lines => {
        r#"void main() {
  var text = """alpha
beta""";
  print(text.split('\n').length);
}"#,
        ["2"]
    };

    triple_quoted_preserves_embedded_single_quotes => {
        r#"void main() {
  var text = '''it's fine''';
  print(text.contains("'"));
}"#,
        ["true"]
    };

    triple_quoted_preserves_embedded_double_quotes => {
        r#"void main() {
  var text = """say "hi""";
  print(text.contains("hi"));
}"#,
        ["true"]
    };

    raw_string_preserves_backslash_n => {
        r#"void main() {
  var s = r'line\nline';
  print(s.length);
}"#,
        ["9"]
    };

    raw_string_does_not_interpolate_dollar_var => {
        r#"void main() {
  var name = 'Bob';
  var s = r'Hi $name';
  print(s);
}"#,
        ["Hi $name"]
    };

    raw_string_preserves_backslashes => {
        r#"void main() {
  var s = r'C:\Users\test';
  print(s.contains('\\'));
}"#,
        ["true"]
    };

    escape_newline_in_single_quoted_string => {
        r#"void main() {
  var s = 'a\nb';
  print(s.split('\n').length);
}"#,
        ["2"]
    };

    escape_tab_in_single_quoted_string => {
        r#"void main() {
  var s = 'a\tb';
  print(s.contains('\t'));
}"#,
        ["true"]
    };

    escape_backslash_in_single_quoted_string => {
        r#"void main() {
  var s = 'a\\b';
  print(s.length);
}"#,
        ["3"]
    };

    escape_single_quote_inside_single_quoted_string => {
        r#"void main() {
  var s = 'it\'s';
  print(s);
}"#,
        ["it's"]
    };

    escape_double_quote_inside_double_quoted_string => {
        r#"void main() {
  var s = "say \"hi\"";
  print(s);
}"#,
        ["say \"hi\""]
    };

    unicode_escape_u0041_is_capital_a => {
        r#"void main() {
  var s = '\u0041';
  print(s);
}"#,
        ["A"]
    };

    hex_escape_x42_is_capital_b => {
        r#"void main() {
  var s = '\x42';
  print(s);
}"#,
        ["B"]
    };

    dollar_sign_without_identifier_is_literal => {
        r#"void main() {
  print('cost is $5');
}"#,
        ["cost is $5"]
    };

    expr_interpolation_with_null_coalescing => {
        r#"void main() {
  String? missing;
  var fallback = 'none';
  print('val=${missing ?? fallback}');
}"#,
        ["val=none"]
    };

    expr_interpolation_list_length => {
        r#"void main() {
  var items = [1, 2, 3, 4];
  print('count=${items.length}');
}"#,
        ["count=4"]
    };

    expr_interpolation_string_concat_inside => {
        r#"void main() {
  var a = 'foo';
  var b = 'bar';
  print('${a + b}');
}"#,
        ["foobar"]
    };

    interpolation_preserves_surrounding_spaces => {
        r#"void main() {
  var n = 1;
  print('a $n b');
}"#,
        ["a 1 b"]
    };

    interpolation_with_bool_true => {
        r#"void main() {
  var ok = true;
  print('ok=$ok');
}"#,
        ["ok=true"]
    };

    interpolation_with_bool_false => {
        r#"void main() {
  var ok = false;
  print('ok=$ok');
}"#,
        ["ok=false"]
    };

    interpolation_with_double_value => {
        r#"void main() {
  var pi = 3.5;
  print('pi=$pi');
}"#,
        ["pi=3.5"]
    };

    interpolation_in_double_quoted_string => {
        r#"void main() {
  var n = 9;
  print("n=$n");
}"#,
        ["n=9"]
    };

    nested_braces_in_expr_interpolation => {
        r#"void main() {
  var tag = 'item';
  print('${'<$tag>'}');
}"#,
        ["<item>"]
    };

    adjacent_strings_with_interpolation_in_middle => {
        r#"void main() {
  var n = 5;
  print('a=' '$n' '!');
}"#,
        ["a=5!"]
    };

    raw_string_with_single_quotes_inside => {
        r#"void main() {
  var s = r"can't break me";
  print(s.length > 5);
}"#,
        ["true"]
    };

    triple_quoted_string_interpolates_dollar_var => {
        r#"void main() {
  var n = 99;
  var s = '''value $n''';
  print(s);
}"#,
        ["value 99"]
    };

    escape_carriage_return_in_string => {
        r#"void main() {
  var s = 'a\rb';
  print(s.length);
}"#,
        ["3"]
    };

    interpolation_after_adjacent_literal_prefix => {
        r#"void main() {
  var id = 42;
  print('id:' '$id');
}"#,
        ["id:42"]
    };
}
