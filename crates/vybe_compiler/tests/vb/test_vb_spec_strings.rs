use super::helpers::run_vb;

macro_rules! vb_expr_spec {
    ($name:ident, $expr:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let src = format!(
                r#"Module M
    Sub Main()
        Console.WriteLine({})
    End Sub
End Module
"#,
                $expr
            );
            let out = run_vb(&src);
            assert_eq!(out, vec![super::helpers::dotnet_expected_one($expected)]);
        }
    };
}

macro_rules! vb_full_spec {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, super::helpers::dotnet_expected_lines(&[$($expected),*]));
        }
    };
}

vb_expr_spec!(
    string_spec_instrrev_returns_last_match_position,
    r#"InStrRev("banana", "na")"#,
    "5"
);
vb_expr_spec!(
    string_spec_instrrev_returns_zero_when_missing,
    r#"InStrRev("banana", "zz")"#,
    "0"
);
vb_expr_spec!(
    string_spec_strcomp_returns_zero_for_equal_values,
    r#"StrComp("alpha", "alpha")"#,
    "0"
);
vb_expr_spec!(
    string_spec_strcomp_binary_compare_distinguishes_case,
    r#"StrComp("Alpha", "alpha", CompareMethod.Binary)"#,
    "-1"
);
vb_expr_spec!(
    string_spec_string_function_repeats_single_character,
    r#"String(5, "*"c)"#,
    "*****"
);
vb_expr_spec!(
    string_spec_string_function_repeats_string_seed,
    r#"String(3, "A"c)"#,
    "AAA"
);
vb_full_spec!(
    string_spec_filter_returns_only_matching_entries,
    r#"Module M
    Sub Main()
        Dim items() As String = {"alpha", "beta", "gamma"}
        Dim filtered() As String = Filter(items, "a")
        Console.WriteLine(Join(filtered, ","))
    End Sub
End Module"#,
    ["alpha,beta,gamma"]
);
vb_full_spec!(
    string_spec_filter_can_exclude_matching_entries,
    r#"Module M
    Sub Main()
        Dim items() As String = {"alpha", "beta", "gamma"}
        Dim filtered() As String = Filter(items, "mm", False)
        Console.WriteLine(Join(filtered, ","))
    End Sub
End Module"#,
    ["alpha,beta"]
);
vb_expr_spec!(
    string_spec_format_pads_numeric_output,
    r#"Format(12, "0000")"#,
    "0012"
);
vb_expr_spec!(
    string_spec_format_renders_currency_pattern,
    r#"Format(12.5, "$0.00")"#,
    "$12.50"
);
vb_expr_spec!(
    string_spec_format_renders_percentage_pattern,
    r#"Format(0.256, "0.0%")"#,
    "25.6%"
);
vb_expr_spec!(
    string_spec_format_renders_short_date_pattern,
    r#"Format(#5/14/2024#, "Short Date")"#,
    "5/14/2024"
);
vb_expr_spec!(
    string_spec_format_renders_short_time_pattern,
    r#"Format(#5/14/2024 3:45 PM#, "Short Time")"#,
    "3:45 PM"
);
vb_full_spec!(
    string_spec_mid_statement_replaces_inner_span,
    r#"Module M
    Sub Main()
        Dim text As String = "abcdef"
        Mid(text, 3, 2) = "XY"
        Console.WriteLine(text)
    End Sub
End Module"#,
    ["abXYef"]
);
vb_full_spec!(
    string_spec_mid_statement_can_overwrite_from_given_position,
    r#"Module M
    Sub Main()
        Dim text As String = "planet"
        Mid(text, 4) = "NET"
        Console.WriteLine(text)
    End Sub
End Module"#,
    ["plaNET"]
);
vb_full_spec!(
    string_spec_lset_left_aligns_fixed_length_target,
    r#"Module M
    Sub Main()
        Dim text As String * 6
        LSet text = "VB"
        Console.WriteLine("[" & text & "]")
    End Sub
End Module"#,
    ["[VB    ]"]
);
vb_full_spec!(
    string_spec_rset_right_aligns_fixed_length_target,
    r#"Module M
    Sub Main()
        Dim text As String * 6
        RSet text = "VB"
        Console.WriteLine("[" & text & "]")
    End Sub
End Module"#,
    ["[    VB]"]
);
vb_expr_spec!(
    string_spec_len_counts_empty_string_as_zero,
    r#"Len("")"#,
    "0"
);
vb_full_spec!(
    string_spec_len_counts_fixed_length_string_buffer,
    r#"Module M
    Sub Main()
        Dim text As String * 4
        text = "AB"
        Console.WriteLine(Len(text))
    End Sub
End Module"#,
    ["4"]
);
vb_expr_spec!(
    string_spec_trim_preserves_internal_spaces,
    r#"Trim("  hello  world  ")"#,
    "hello  world"
);
vb_expr_spec!(
    string_spec_instr_with_start_position_skips_early_matches,
    r#"InStr(3, "banana", "na")"#,
    "5"
);
vb_expr_spec!(
    string_spec_instrrev_with_start_position_limits_right_search,
    r#"InStrRev("banana", "na", 4)"#,
    "3"
);
vb_expr_spec!(
    string_spec_replace_with_start_parameter_skips_prefix,
    r#"Replace("banana", "na", "XY", 3)"#,
    "banXYna"
);
vb_expr_spec!(
    string_spec_replace_with_count_parameter_limits_changes,
    r#"Replace("banana", "na", "XY", 1, 1)"#,
    "baXYna"
);
vb_expr_spec!(
    string_spec_join_rebuilds_csv_row_with_semicolon,
    r#"Join(Split("a,b,c", ","), ";")"#,
    "a;b;c"
);
vb_expr_spec!(
    string_spec_split_with_count_argument_limits_segments,
    r#"UBound(Split("a,b,c,d", ",", 3))"#,
    "2"
);
vb_expr_spec!(
    string_spec_split_with_compare_argument_uses_text_match,
    r#"UBound(Split("A|a|B", "a", -1, CompareMethod.Text))"#,
    "2"
);
vb_expr_spec!(
    string_spec_ascw_returns_unicode_code_point,
    r#"AscW("é")"#,
    "233"
);
vb_expr_spec!(
    string_spec_chrw_returns_unicode_character,
    r#"ChrW(9731)"#,
    "☃"
);
vb_expr_spec!(
    string_spec_vbtab_concatenates_with_labels,
    r#""left" & vbTab & "right""#,
    "left\tright"
);
vb_expr_spec!(
    string_spec_vbcrlf_separates_two_console_lines,
    r#""top" & vbCrLf & "bottom""#,
    "top\r\nbottom"
);
vb_expr_spec!(
    string_spec_vbnewline_can_join_header_and_body,
    r#""header" & vbNewLine & "body""#,
    "header\r\nbody"
);
vb_expr_spec!(
    string_spec_left_returns_whole_string_when_length_is_large,
    r#"Left("cat", 10)"#,
    "cat"
);
vb_expr_spec!(
    string_spec_right_returns_whole_string_when_length_is_large,
    r#"Right("cat", 10)"#,
    "cat"
);
vb_expr_spec!(
    string_spec_mid_two_argument_returns_tail_section,
    r#"Mid("vibecode", 5)"#,
    "code"
);
vb_expr_spec!(
    string_spec_strreverse_reverses_letters_only,
    r#"StrReverse("drawer")"#,
    "reward"
);
vb_expr_spec!(
    string_spec_space_can_build_padding_inside_brackets,
    r#""[" & Space(3) & "]""#,
    "[   ]"
);
vb_expr_spec!(
    string_spec_isnumeric_accepts_signed_decimal_string,
    r#"IsNumeric("-12.5")"#,
    "true"
);
vb_expr_spec!(
    string_spec_isnumeric_rejects_alphabetic_word,
    r#"IsNumeric("hello")"#,
    "false"
);
vb_expr_spec!(
    string_spec_cstr_preserves_leading_minus_sign,
    r#"CStr(-42)"#,
    "-42"
);
vb_expr_spec!(
    string_spec_val_stops_parsing_at_first_non_numeric_character,
    r#"Val("12.5kg")"#,
    "12.5"
);
vb_expr_spec!(
    string_spec_join_can_use_empty_delimiter,
    r#"Join(Split("v-b", "-"), "")"#,
    "vb"
);
vb_expr_spec!(
    string_spec_split_can_preserve_empty_items_between_delimiters,
    r#"UBound(Split("a,,b", ","))"#,
    "2"
);
vb_expr_spec!(
    string_spec_filter_can_match_substrings_inside_words,
    r#"UBound(Filter(Array("stone", "tone", "ring"), "one"))"#,
    "1"
);
vb_expr_spec!(
    string_spec_filter_can_keep_case_variant_matches_in_text_mode,
    r#"UBound(Filter(Array("Alpha", "beta", "ALP"), "alp", True, CompareMethod.Text))"#,
    "1"
);
vb_expr_spec!(
    string_spec_ucase_transforms_mixed_case_sentence,
    r#"UCase("Visual Basic")"#,
    "VISUAL BASIC"
);
vb_expr_spec!(
    string_spec_lcase_transforms_mixed_case_sentence,
    r#"LCase("Visual Basic")"#,
    "visual basic"
);
vb_expr_spec!(
    string_spec_replace_can_remove_target_text_by_using_empty_replacement,
    r#"Replace("abracadabra", "a", "")"#,
    "brcdbr"
);
vb_expr_spec!(
    string_spec_instr_returns_one_based_index_for_first_match,
    r#"InStr("hello", "e")"#,
    "2"
);
vb_expr_spec!(
    string_spec_strcomp_text_compare_ignores_case,
    r#"StrComp("Visual", "visual", CompareMethod.Text)"#,
    "0"
);
