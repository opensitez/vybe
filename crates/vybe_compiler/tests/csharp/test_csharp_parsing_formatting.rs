use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(int_parse_converts_decimal_text_to_number, r#"Console.WriteLine(int.Parse("42") + 1);"#, ["43"]);
csharp_case!(int_try_parse_reports_true_for_valid_digits, r#"var ok = int.TryParse("42", out var value); Console.WriteLine(ok); Console.WriteLine(value);"#, ["True", "42"]);
csharp_case!(int_try_parse_reports_false_for_invalid_text, r#"var ok = int.TryParse("4x", out var value); Console.WriteLine(ok); Console.WriteLine(value);"#, ["False", "0"]);
csharp_case!(double_parse_reads_fractional_value, r#"Console.WriteLine(double.Parse("3.5") + 0.5);"#, ["4"]);
csharp_case!(decimal_parse_reads_decimal_literal_text, r#"Console.WriteLine(decimal.Parse("7.25"));"#, ["7.25"]);
csharp_case!(bool_parse_reads_true_text, r#"Console.WriteLine(bool.Parse("True"));"#, ["True"]);
csharp_case!(char_parse_reads_single_character_text, r#"Console.WriteLine(char.Parse("Z"));"#, ["Z"]);
csharp_case!(byte_parse_reads_small_integer_text, r#"Console.WriteLine(byte.Parse("12") + 1);"#, ["13"]);
csharp_case!(long_parse_reads_large_integer_text, r#"Console.WriteLine(long.Parse("123456") + 1);"#, ["123457"]);
csharp_case!(float_parse_reads_fractional_text, r#"Console.WriteLine(float.Parse("2.5") + 0.5f);"#, ["3"]);
csharp_case!(string_format_replaces_indexed_placeholders, r#"Console.WriteLine(string.Format("{0}-{1}", "A", 3));"#, ["A-3"]);
csharp_case!(to_string_decimal_format_pads_with_leading_zeroes, r#"Console.WriteLine(7.ToString("D4"));"#, ["0007"]);
csharp_case!(to_string_hex_format_outputs_uppercase_hex_digits, r#"Console.WriteLine(255.ToString("X"));"#, ["FF"]);
csharp_case!(interpolated_string_embeds_computed_values, r#"Console.WriteLine($"sum={2 + 3}");"#, ["sum=5"]);
csharp_case!(number_format_can_render_fixed_point_precision, r#"Console.WriteLine(3.14159.ToString("F2"));"#, ["3.14"]);
csharp_case!(number_format_can_render_percentage_output, r#"Console.WriteLine(0.25.ToString("P0"));"#, ["25 %"]);
csharp_case!(parse_signed_integer_text_preserves_negative_sign, r#"Console.WriteLine(int.Parse("-9"));"#, ["-9"]);
csharp_case!(trim_then_parse_allows_surrounding_whitespace, r#"Console.WriteLine(int.Parse(" 12 ".Trim()));"#, ["12"]);
csharp_case!(convert_to_string_can_render_integer_value, r#"Console.WriteLine(System.Convert.ToString(25));"#, ["25"]);
csharp_case!(string_join_renders_delimited_sequence, r#"Console.WriteLine(string.Join("|", new[] { "a", "b", "c" }));"#, ["a|b|c"]);