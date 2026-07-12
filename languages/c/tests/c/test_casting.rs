use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    cast_int_to_double_preserves_value => { declarations: "", body: "printf(\"%.1f\\n\", (double)7);\nreturn 0;", expect: ["7.0"] },
    cast_double_to_int_truncates_fraction => { declarations: "", body: "printf(\"%d\\n\", (int)7.9);\nreturn 0;", expect: ["7"] },
    cast_negative_double_to_int_truncates_toward_zero => { declarations: "", body: "printf(\"%d\\n\", (int)-7.9);\nreturn 0;", expect: ["-7"] },
    cast_char_to_int_uses_character_code => { declarations: "", body: "printf(\"%d\\n\", (int)'A');\nreturn 0;", expect: ["65"] },
    cast_int_to_char_uses_character_code => { declarations: "", body: "printf(\"%c\\n\", (char)66);\nreturn 0;", expect: ["B"] },
    cast_before_division_changes_result_type => { declarations: "", body: "printf(\"%.2f\\n\", (double)1 / 2);\nreturn 0;", expect: ["0.50"] },
    cast_after_integer_division_cannot_restore_fraction => { declarations: "", body: "printf(\"%.2f\\n\", (double)(1 / 2));\nreturn 0;", expect: ["0.00"] },
    cast_sum_to_char_after_integer_addition => { declarations: "", body: "printf(\"%c\\n\", (char)('A' + 2));\nreturn 0;", expect: ["C"] },
    cast_float_literal_to_int_inside_expression => { declarations: "", body: "printf(\"%d\\n\", (int)3.9 + 1);\nreturn 0;", expect: ["4"] },
    cast_parenthesized_sum_to_int_after_addition => { declarations: "", body: "printf(\"%d\\n\", (int)(3.9 + 1.2));\nreturn 0;", expect: ["5"] },
    cast_negative_int_to_unsigned_wraps_modulo_range => { declarations: "", body: "printf(\"%u\\n\", (unsigned int)-1);\nreturn 0;", expect: ["4294967295"] },
    cast_unsigned_back_to_signed_changes_display => { declarations: "unsigned int x = (unsigned int)-1;", body: "printf(\"%d\\n\", (int)x);\nreturn 0;", expect: ["-1"] },
    cast_comparison_result_to_double => { declarations: "", body: "printf(\"%.1f\\n\", (double)(3 > 2));\nreturn 0;", expect: ["1.0"] },
    cast_char_result_back_to_int_after_increment => { declarations: "char c = 'a';", body: "c = (char)(c + 1);\nprintf(\"%d\\n\", (int)c);\nreturn 0;", expect: ["98"] },
    cast_to_short_truncates_large_value_bits => { declarations: "", body: "printf(\"%d\\n\", (short)65537);\nreturn 0;", expect: ["1"] },
    cast_to_long_preserves_integer_value => { declarations: "", body: "printf(\"%ld\\n\", (long)42);\nreturn 0;", expect: ["42"] },
    cast_float_expression_to_int_after_multiplication => { declarations: "", body: "printf(\"%d\\n\", (int)(2.5 * 2.0));\nreturn 0;", expect: ["5"] },
    cast_int_expression_to_double_for_formatting => { declarations: "", body: "int x = 7;\nprintf(\"%.1f\\n\", (double)x);\nreturn 0;", expect: ["7.0"] },
    cast_zero_to_pointer_can_compare_with_null => { declarations: "int *p = (int *)0;", body: "if (p == NULL) puts(\"null\"); else puts(\"bad\");\nreturn 0;", expect: ["null"] },
    cast_pointer_to_void_and_back_preserves_address => { declarations: "int x = 4; int *p = &x; void *q = (void *)p;", body: "printf(\"%d\\n\", *(int *)q);\nreturn 0;", expect: ["4"] },
    cast_double_to_int_inside_comparison => { declarations: "", body: "printf(\"%d\\n\", ((int)3.9) == 3);\nreturn 0;", expect: ["1"] },
    cast_char_pointer_from_string_literal_reads_character => { declarations: "char *p = (char *)\"vybe\";", body: "printf(\"%c\\n\", p[1]);\nreturn 0;", expect: ["y"] },
    nested_casts_can_round_trip_small_integer => { declarations: "", body: "printf(\"%d\\n\", (int)(double)(char)65);\nreturn 0;", expect: ["65"] },
    cast_in_ternary_preserves_selected_type => { declarations: "", body: "printf(\"%.1f\\n\", 1 ? (double)2 : 3);\nreturn 0;", expect: ["2.0"] },
    cast_before_modulo_expression_is_integer_based => { declarations: "", body: "printf(\"%d\\n\", (int)7.9 % 3);\nreturn 0;", expect: ["1"] },
    cast_array_element_to_double_for_division => { declarations: "int arr[2] = {1, 2};", body: "printf(\"%.2f\\n\", (double)arr[0] / arr[1]);\nreturn 0;", expect: ["0.50"] },
    cast_double_sum_to_char_yields_ascii_character => { declarations: "", body: "printf(\"%c\\n\", (char)(65.0 + 1.0));\nreturn 0;", expect: ["B"] },
    cast_boolean_like_expression_to_char_displays_code => { declarations: "", body: "printf(\"%d\\n\", (char)(3 > 2));\nreturn 0;", expect: ["1"] },
    cast_negative_char_code_to_unsigned_char_wraps => { declarations: "", body: "printf(\"%u\\n\", (unsigned char)-1);\nreturn 0;", expect: ["255"] },
    cast_pointer_difference_to_int_keeps_offset => { declarations: "int arr[3] = {1, 2, 3};", body: "printf(\"%d\\n\", (int)(&arr[2] - &arr[0]));\nreturn 0;", expect: ["2"] }
}
