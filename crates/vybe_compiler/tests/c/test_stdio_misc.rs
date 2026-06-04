use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<string.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    puts_prints_literal_line => { declarations: "", body: "puts(\"hello\");\nreturn 0;", expect: ["hello"] },
    puts_prints_string_variable => { declarations: "char *text = \"vybe\";", body: "puts(text);\nreturn 0;", expect: ["vybe"] },
    putchar_result_can_be_formatted_as_character => { declarations: "", body: "printf(\"%c\\n\", putchar('A'));\nreturn 0;", expect: ["A"] },
    putchar_result_can_be_formatted_as_decimal_code => { declarations: "", body: "printf(\"%d\\n\", putchar('B'));\nreturn 0;", expect: ["66"] },
    printf_can_emit_three_integers => { declarations: "", body: "printf(\"%d %d %d\\n\", 1, 2, 3);\nreturn 0;", expect: ["1 2 3"] },
    printf_can_mix_string_and_float => { declarations: "", body: "printf(\"%s %.1f\\n\", \"pi\", 3.1);\nreturn 0;", expect: ["pi 3.1"] },
    sprintf_can_fill_buffer_for_puts => { declarations: "char buf[64];", body: "sprintf(buf, \"%d-%s\", 7, \"cat\");\nputs(buf);\nreturn 0;", expect: ["7-cat"] },
    sprintf_can_format_float_into_buffer => { declarations: "char buf[64];", body: "sprintf(buf, \"%.2f\", 2.5);\nputs(buf);\nreturn 0;", expect: ["2.50"] },
    snprintf_can_limit_visible_text => { declarations: "char buf[64];", body: "snprintf(buf, 4, \"%s\", \"alphabet\");\nputs(buf);\nreturn 0;", expect: ["alp"] },
    fprintf_to_stdout_can_forward_integer_and_string => { declarations: "", body: "fprintf(stdout, \"%d %s\\n\", 4, \"fish\");\nreturn 0;", expect: ["4 fish"] },
    printf_can_emit_percent_literal_between_values => { declarations: "", body: "printf(\"%d%% %s\\n\", 50, \"done\");\nreturn 0;", expect: ["50% done"] },
    printf_can_emit_character_from_integer_expression => { declarations: "", body: "printf(\"%c\\n\", 65 + 1);\nreturn 0;", expect: ["B"] },
    puts_can_print_empty_string => { declarations: "", body: "puts(\"\");\nreturn 0;", expect: [""] },
    printf_can_emit_zero_with_padding => { declarations: "", body: "printf(\"%04d\\n\", 0);\nreturn 0;", expect: ["0000"] },
    sprintf_return_value_can_feed_puts => { declarations: "char buf[64];", body: "puts(sprintf(buf, \"%s:%d\", \"id\", 9));\nreturn 0;", expect: ["id:9"] },
    printf_can_emit_negative_float_with_precision => { declarations: "", body: "printf(\"%.2f\\n\", -3.5);\nreturn 0;", expect: ["-3.50"] },
    fprintf_can_emit_character => { declarations: "", body: "fprintf(stdout, \"%c\\n\", 'Z');\nreturn 0;", expect: ["Z"] },
    snprintf_zero_length_payload_keeps_empty_visible_buffer => { declarations: "char buf[8] = \"seed\";", body: "snprintf(buf, 1, \"%s\", \"abc\");\nputs(buf);\nreturn 0;", expect: [""] },
    sprintf_can_write_adjacent_bracketed_values => { declarations: "char buf[64];", body: "sprintf(buf, \"[%d][%d]\", 1, 2);\nputs(buf);\nreturn 0;", expect: ["[1][2]"] },
    printf_can_emit_hex_and_decimal_same_call => { declarations: "", body: "printf(\"%x %d\\n\", 255, 255);\nreturn 0;", expect: ["ff 255"] },
    printf_can_emit_octal_and_unsigned_same_call => { declarations: "", body: "printf(\"%o %u\\n\", 8, 8u);\nreturn 0;", expect: ["10 8"] },
    puts_after_sprintf_can_show_signed_integer => { declarations: "char buf[64];", body: "sprintf(buf, \"%+d\", 7);\nputs(buf);\nreturn 0;", expect: ["+7"] },
    printf_can_use_width_on_string => { declarations: "", body: "printf(\"%5s\\n\", \"go\");\nreturn 0;", expect: ["   go"] },
    sprintf_can_capture_scientific_notation => { declarations: "char buf[64];", body: "sprintf(buf, \"%.1e\", 1200.0);\nputs(buf);\nreturn 0;", expect: ["1.2e+3"] },
    fprintf_can_forward_mixed_arguments => { declarations: "", body: "fprintf(stdout, \"%s %d %c\\n\", \"mix\", 4, 'Q');\nreturn 0;", expect: ["mix 4 Q"] },
    printf_can_emit_multiple_lines_from_one_format => { declarations: "", body: "printf(\"a\\nb\\n\");\nreturn 0;", expect: ["a", "b"] },
    sprintf_can_reuse_same_buffer_variable => { declarations: "char buf[64];", body: "sprintf(buf, \"%d\", 3);\nputs(buf);\nsprintf(buf, \"%d\", 4);\nputs(buf);\nreturn 0;", expect: ["3", "4"] },
    fprintf_inside_loop_can_emit_each_iteration => { declarations: "", body: "for (int i = 0; i < 2; i++) fprintf(stdout, \"%d\\n\", i);\nreturn 0;", expect: ["0", "1"] },
    puts_can_print_buffer_written_by_snprintf => { declarations: "char buf[64];", body: "snprintf(buf, 64, \"%s-%d\", \"x\", 3);\nputs(buf);\nreturn 0;", expect: ["x-3"] },
    printf_can_emit_char_array_as_string => { declarations: "char text[] = \"abc\";", body: "printf(\"%s\\n\", text);\nreturn 0;", expect: ["abc"] },
    sprintf_can_format_octal_into_buffer => { declarations: "char buf[64];", body: "sprintf(buf, \"%o\", 9);\nputs(buf);\nreturn 0;", expect: ["11"] },
    printf_can_emit_upper_hex => { declarations: "", body: "printf(\"%X\\n\", 255);\nreturn 0;", expect: ["FF"] },
    fprintf_can_emit_percent_literal => { declarations: "", body: "fprintf(stdout, \"%% done\\n\");\nreturn 0;", expect: ["% done"] },
    puts_can_print_result_of_strcat => { declarations: "char left[16] = \"ab\";", body: "puts(strcat(left, \"cd\"));\nreturn 0;", expect: ["abcd"] },
    printf_can_emit_signed_zero_with_plus_flag => { declarations: "", body: "printf(\"%+d\\n\", 0);\nreturn 0;", expect: ["+0"] },
    sprintf_can_format_two_floats => { declarations: "char buf[64];", body: "sprintf(buf, \"%.1f %.1f\", 1.5, 2.5);\nputs(buf);\nreturn 0;", expect: ["1.5 2.5"] },
    snprintf_empty_format_can_clear_buffer => { declarations: "char buf[8] = \"seed\";", body: "snprintf(buf, 8, \"\");\nputs(buf);\nreturn 0;", expect: [""] },
    printf_can_emit_pointer_difference_as_decimal => { declarations: "int values[3] = {1, 2, 3};", body: "printf(\"%d\\n\", (int)(&values[2] - &values[0]));\nreturn 0;", expect: ["2"] },
    putchar_result_can_participate_in_arithmetic => { declarations: "", body: "printf(\"%d\\n\", putchar('A') + 1);\nreturn 0;", expect: ["66"] },
    fprintf_can_emit_integer_with_width => { declarations: "", body: "fprintf(stdout, \"%4d\\n\", 7);\nreturn 0;", expect: ["   7"] }
}
