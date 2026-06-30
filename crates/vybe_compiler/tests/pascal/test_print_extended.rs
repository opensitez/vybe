/// Write/WriteLn formatting with widths and edge cases.
use super::helpers::run_pascal;

#[test]
fn write_integer_zero_width() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(42:0); WriteLn(''); end."#
        ),
        &["42"]
    );
}

#[test]
fn write_integer_width_five_pad() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(7:5); WriteLn(''); end."#
        ),
        &["    7"]
    );
}

#[test]
fn write_integer_width_three_exact() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(123:3); WriteLn(''); end."#
        ),
        &["123"]
    );
}

#[test]
fn write_integer_width_exceeds_value() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(12345:3); WriteLn(''); end."#
        ),
        &["12345"]
    );
}

#[test]
fn write_negative_integer_width() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(-42:6); WriteLn(''); end."#
        ),
        &["   -42"]
    );
}

#[test]
fn writeln_integer_with_width() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(99:4); end."#
        ),
        &["  99"]
    );
}

#[test]
fn write_string_width_pad_right() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write('ab':5); WriteLn(''); end."#
        ),
        &["ab   "]
    );
}

#[test]
fn write_string_width_truncates_display() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write('hello':3); WriteLn(''); end."#
        ),
        &["hello"]
    );
}

#[test]
fn write_real_two_decimals() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(3.14159:0:2); WriteLn(''); end."#
        ),
        &["3.14"]
    );
}

#[test]
fn write_real_width_and_decimals() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(2.5:8:1); WriteLn(''); end."#
        ),
        &["     2.5"]
    );
}

#[test]
fn writeln_real_scientific_style() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(1000.0:10:2); end."#
        ),
        &["   1000.00"]
    );
}

#[test]
fn write_multiple_width_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(1:3); Write(2:3); WriteLn(3:3); end."#
        ),
        &["  1  2  3"]
    );
}

#[test]
fn write_concat_after_width_field() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(5:4); Write('x'); WriteLn(''); end."#
        ),
        &["   5x"]
    );
}

#[test]
fn writeln_empty_between_width_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(1:2); WriteLn(2:2); end."#
        ),
        &[" 1", " 2"]
    );
}

#[test]
fn write_zero_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(0:4); WriteLn(''); end."#
        ),
        &["   0"]
    );
}

#[test]
fn write_boolean_true_default() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(True); end."#
        ),
        &["True"]
    );
}

#[test]
fn write_boolean_false_default() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(False); end."#
        ),
        &["False"]
    );
}

#[test]
fn write_char_literal_direct() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn('Z'); end."#
        ),
        &["Z"]
    );
}

#[test]
fn write_hex_via_format_style() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(IntToHex(255, 4)); end."#
        ),
        &["00FF"]
    );
}

#[test]
fn write_loop_index_aligned() {
    assert_eq!(
        run_pascal(
            r#"program T; var i: Integer; begin for i := 1 to 3 do WriteLn(i:3); end."#
        ),
        &["  1", "  2", "  3"]
    );
}

#[test]
fn write_table_header_separator() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn('Name':10); WriteLn(StringOfChar('-', 10)); end."#
        ),
        &["Name      ", "----------"]
    );
}

#[test]
fn write_mixed_string_int_real() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write('n='); Write(7); Write(' r='); WriteLn(1.5:4:1); end."#
        ),
        &["n=7 r= 1.5"]
    );
}

#[test]
fn write_partial_line_then_writeln() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write('a'); Write('b'); WriteLn('c'); end."#
        ),
        &["abc"]
    );
}

#[test]
fn write_only_no_trailing_newline() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write('no'); end."#
        ),
        &["no"]
    );
}

#[test]
fn writeln_after_write_completes_line() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write('pre'); WriteLn('post'); end."#
        ),
        &["prepost"]
    );
}

#[test]
fn write_integer_large_value() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(1000000); end."#
        ),
        &["1000000"]
    );
}

#[test]
fn write_real_negative_value() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(-1.25:6:2); WriteLn(''); end."#
        ),
        &[" -1.25"]
    );
}

#[test]
fn write_string_empty() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(''); end."#
        ),
        &[""]
    );
}

#[test]
fn write_format_percent_in_string() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn('100%'); end."#
        ),
        &["100%"]
    );
}

#[test]
fn write_nested_expression_in_width() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn((3 + 4):5); end."#
        ),
        &["    7"]
    );
}

#[test]
fn write_function_result_with_width() {
    assert_eq!(
        run_pascal(
            r#"program T; function N: Integer; begin Result := 88; end; begin WriteLn(N:4); end."#
        ),
        &["  88"]
    );
}

#[test]
fn write_record_field_via_with() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR = record V: Integer; end; var r: TR; begin r.V := 15; with r do WriteLn(V:4); end."#
        ),
        &["  15"]
    );
}

#[test]
fn write_array_element_formatted() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..2] of Integer; begin a[0]:=1; a[1]:=2; a[2]:=3; WriteLn(a[1]:3); end."#
        ),
        &["  2"]
    );
}

#[test]
fn write_columns_in_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Col(a, b: Integer); begin Write(a:4); WriteLn(b:4); end; begin Col(1, 20); Col(300, 4); end."#
        ),
        &["   1  20", " 300   4"]
    );
}

#[test]
fn write_real_zero_decimals() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(5.0:0:0); end."#
        ),
        &["5"]
    );
}

#[test]
fn write_integer_min_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; begin n := -1; WriteLn(n); end."#
        ),
        &["-1"]
    );
}

#[test]
fn write_sequential_writes_same_line() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(1); Write(2); Write(3); WriteLn(4); end."#
        ),
        &["1234"]
    );
}

#[test]
fn write_width_one_minimal_pad() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write(9:1); WriteLn(''); end."#
        ),
        &["9"]
    );
}

#[test]
fn writeln_multiple_args_default() {
    assert_eq!(
        run_pascal(
            r#"program T; begin WriteLn(1, 2, 3); end."#
        ),
        &["1 2 3"]
    );
}

#[test]
fn write_currency_style_real() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write('$'); WriteLn(19.99:0:2); end."#
        ),
        &["$19.99"]
    );
}

#[test]
fn write_padding_with_spaces_between() {
    assert_eq!(
        run_pascal(
            r#"program T; begin Write('a':3); Write('|'); Write('b':3); WriteLn(''); end."#
        ),
        &["a  |b  "]
    );
}
