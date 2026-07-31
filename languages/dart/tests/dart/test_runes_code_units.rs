//! String.codeUnitAt, codeUnits, runes, fromCharCodes, surrogate pair lengths.

dart_cases! {
    code_unit_at_first_ascii => {
        r#"void main() {
  print('ABC'.codeUnitAt(0));
}"#,
        ["65"]
    };

    code_unit_at_second_ascii => {
        r#"void main() {
  print('ABC'.codeUnitAt(1));
}"#,
        ["66"]
    };

    code_unit_at_last_index => {
        r#"void main() {
  print('ABC'.codeUnitAt(2));
}"#,
        ["67"]
    };

    code_units_list_ascii => {
        r#"void main() {
  print('Hi'.codeUnits);
}"#,
        ["[72, 105]"]
    };

    code_units_length_matches_string_length => {
        r#"void main() {
  var s = 'dart';
  print(s.codeUnits.length == s.length);
}"#,
        ["true"]
    };

    from_char_codes_ascii => {
        r#"void main() {
  print(String.fromCharCodes([65, 66, 67]));
}"#,
        ["ABC"]
    };

    from_char_codes_empty_list => {
        r#"void main() {
  print(String.fromCharCodes([]));
}"#,
        [""]
    };

    from_char_codes_single => {
        r#"void main() {
  print(String.fromCharCodes([90]));
}"#,
        ["Z"]
    };

    from_char_codes_round_trip => {
        r#"void main() {
  var codes = 'hello'.codeUnits;
  print(String.fromCharCodes(codes));
}"#,
        ["hello"]
    };

    runes_length_ascii_equals_length => {
        r#"void main() {
  var s = 'test';
  print(s.runes.length);
}"#,
        ["4"]
    };

    runes_to_list_ascii => {
        r#"void main() {
  print('ab'.runes.toList());
}"#,
        ["[97, 98]"]
    };

    runes_first_element => {
        r#"void main() {
  print('xy'.runes.first);
}"#,
        ["120"]
    };

    runes_last_element => {
        r#"void main() {
  print('xy'.runes.last);
}"#,
        ["121"]
    };

    emoji_string_length_two_code_units => {
        r#"void main() {
  print('🙂'.length);
}"#,
        ["2"]
    };

    emoji_runes_length_one => {
        r#"void main() {
  print('🙂'.runes.length);
}"#,
        ["1"]
    };

    emoji_code_units_are_surrogate_pair => {
        r#"void main() {
  print('🙂'.codeUnits);
}"#,
        ["[55357, 56842]"]
    };

    emoji_runes_single_code_point => {
        r#"void main() {
  print('🙂'.runes.toList());
}"#,
        ["[128578]"]
    };

    from_char_codes_surrogate_pair_emoji => {
        r#"void main() {
  print(String.fromCharCodes([55357, 56842]));
}"#,
        ["🙂"]
    };

    surrogate_pair_runes_length_vs_length => {
        r#"void main() {
  var s = String.fromCharCodes([55357, 56842]);
  print(s.length);
  print(s.runes.length);
}"#,
        ["2", "1"]
    };

    accented_e_code_units => {
        r#"void main() {
  print('é'.codeUnits);
}"#,
        ["[233]"]
    };

    accented_e_runes_length => {
        r#"void main() {
  print('é'.runes.length);
}"#,
        ["1"]
    };

    cafe_runes_length => {
        r#"void main() {
  print('café'.runes.length);
}"#,
        ["4"]
    };

    cafe_code_units_length => {
        r#"void main() {
  print('café'.codeUnits.length);
}"#,
        ["5"]
    };

    code_unit_at_accented_letter => {
        r#"void main() {
  print('café'.codeUnitAt(3));
}"#,
        ["233"]
    };

    runes_contains_ascii_range => {
        r#"void main() {
  var r = 'A'.runes.first;
  print(r >= 65 && r <= 90);
}"#,
        ["true"]
    };

    string_from_char_codes_newline => {
        r#"void main() {
  print(String.fromCharCodes([10]));
}"#,
        ["\n"]
    };

    string_from_char_codes_tab => {
        r#"void main() {
  print(String.fromCharCodes([9]));
}"#,
        ["\t"]
    };

    code_units_of_space => {
        r#"void main() {
  print(' '.codeUnits);
}"#,
        ["[32]"]
    };

    runes_of_digit_zero => {
        r#"void main() {
  print('0'.runes.first);
}"#,
        ["48"]
    };

    two_emoji_runes_length => {
        r#"void main() {
  print('🙂🙂'.runes.length);
}"#,
        ["2"]
    };

    two_emoji_string_length => {
        r#"void main() {
  print('🙂🙂'.length);
}"#,
        ["4"]
    };

    runes_map_double_ascii => {
        r#"void main() {
  print('ab'.runes.map((r) => r + 1).toList());
}"#,
        ["[98, 99]"]
    };

    code_unit_at_zero_index_empty_error_caught => {
        r#"void main() {
  try {
    print(''.codeUnitAt(0));
  } catch (e) {
    print('caught');
  }
}"#,
        ["caught"]
    };

    from_char_codes_high_bmp => {
        r#"void main() {
  print(String.fromCharCodes([8364]));
}"#,
        ["€"]
    };

    euro_sign_code_units => {
        r#"void main() {
  print('€'.codeUnits);
}"#,
        ["[8364]"]
    };

    euro_sign_runes_length => {
        r#"void main() {
  print('€'.runes.length);
}"#,
        ["1"]
    };

    runes_join_empty => {
        r#"void main() {
  print(''.runes.isEmpty);
}"#,
        ["true"]
    };

    code_units_is_not_empty_for_ascii => {
        r#"void main() {
  print('x'.codeUnits.isNotEmpty);
}"#,
        ["true"]
    };

    from_char_codes_mixed_ascii_and_accent => {
        r#"void main() {
  print(String.fromCharCodes([97, 233]));
}"#,
        ["aé"]
    };

    string_length_vs_runes_mixed_ascii_emoji => {
        r#"void main() {
  var s = 'a🙂b';
  print(s.length);
  print(s.runes.length);
}"#,
        ["4", "3"]
    };

    code_unit_at_emoji_first_surrogate => {
        r#"void main() {
  print('🙂'.codeUnitAt(0));
}"#,
        ["55357"]
    };

    code_unit_at_emoji_second_surrogate => {
        r#"void main() {
  print('🙂'.codeUnitAt(1));
}"#,
        ["56842"]
    };

    runes_string_from_rune_values => {
        r#"void main() {
  print(String.fromCharCodes([128578]));
}"#,
        ["😀"]
    };

    ascii_code_unit_matches_rune => {
        r#"void main() {
  var s = 'Q';
  print(s.codeUnitAt(0) == s.runes.first);
}"#,
        ["true"]
    };

    from_char_codes_zero => {
        r#"void main() {
  print(String.fromCharCodes([0]).codeUnitAt(0));
}"#,
        ["0"]
    };

    runes_where_filter => {
        r#"void main() {
  print('abc'.runes.where((r) => r > 97).length);
}"#,
        ["2"]
    };

    code_units_sublist => {
        r#"void main() {
  print('hello'.codeUnits.sublist(1, 4));
}"#,
        ["[101, 108, 108]"]
    };

    string_from_char_codes_iterable => {
        r#"void main() {
  print(String.fromCharCodes([72, 73].map((c) => c)));
}"#,
        ["HI"]
    };

    flag_emoji_runes_length => {
        r#"void main() {
  print('🇺🇸'.runes.length);
}"#,
        ["2"]
    };

    flag_emoji_string_length => {
        r#"void main() {
  print('🇺🇸'.length);
}"#,
        ["4"]
    };

    // A non-BMP character ANYWHERE in the source used to panic the compiler
    // before it ran: `normalize_parenthesized_is_ternary` scans BYTES and then
    // sliced `&source[j..j + 2]`, which lands inside a 4-byte character.
    // `print('😀');` on its own was enough. Values from real `dart`.
    non_bmp_literal_in_source_compiles => {
        r#"void main() {
  print('😀');
}"#,
        ["😀"]
    };

    non_bmp_length_is_utf16_units => {
        r#"void main() {
  print('😀'.length);
}"#,
        ["2"]
    };

    // The `is`-ternary normalisation that the byte scan exists FOR must still
    // fire in a source that also contains a non-BMP character.
    is_ternary_still_normalised_beside_non_bmp => {
        r#"void main() {
  Object o = 5;
  print((o is int) ? 'a😀b' : 'no');
}"#,
        ["a😀b"]
    };
}
