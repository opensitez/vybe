use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<stdlib.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    atoi_parses_leading_plus_sign => { body: "printf(\"%d\\n\", atoi(\"+17\"));\nreturn 0;", expect: ["17"] },
    atoi_stops_at_first_non_digit => { body: "printf(\"%d\\n\", atoi(\"15cats\"));\nreturn 0;", expect: ["15"] },
    atoi_returns_zero_for_non_numeric_prefix => { body: "printf(\"%d\\n\", atoi(\"cats15\"));\nreturn 0;", expect: ["0"] },
    atol_parses_negative_text => { body: "printf(\"%d\\n\", atol(\"-1234\"));\nreturn 0;", expect: ["-1234"] },
    atof_parses_scientific_notation => { body: "printf(\"%.1f\\n\", atof(\"1.5e2\"));\nreturn 0;", expect: ["150.0"] },
    atof_parses_leading_whitespace => { body: "printf(\"%.1f\\n\", atof(\"  2.5\"));\nreturn 0;", expect: ["2.5"] },
    abs_of_zero_is_zero => { body: "printf(\"%d\\n\", abs(0));\nreturn 0;", expect: ["0"] },
    labs_of_zero_is_zero => { body: "printf(\"%d\\n\", labs(0));\nreturn 0;", expect: ["0"] },
    rand_result_is_within_zero_and_rand_max => { body: "int value = rand(); printf(\"%d\\n\", value >= 0 && value <= RAND_MAX); return 0;", expect: ["1"] },
    rand_can_be_called_twice_and_stay_in_range => { body: "int a = rand(); int b = rand(); printf(\"%d\\n\", a >= 0 && a <= RAND_MAX && b >= 0 && b <= RAND_MAX); return 0;", expect: ["1"] },
    rand_max_constant_is_positive => { body: "printf(\"%d\\n\", RAND_MAX > 0);\nreturn 0;", expect: ["1"] },
    null_macro_compares_equal_to_zero => { body: "printf(\"%d\\n\", NULL == 0);\nreturn 0;", expect: ["1"] },
    malloc_result_can_be_null_checked => { body: "void *p = malloc(16); printf(\"%d\\n\", p != NULL); free(p); return 0;", expect: ["1"] },
    calloc_result_can_be_null_checked => { body: "void *p = calloc(4, 4); printf(\"%d\\n\", p != NULL); free(p); return 0;", expect: ["1"] },
    malloc_memory_can_be_written_through_int_pointer => { body: "int *p = (int *)malloc(sizeof(int)); *p = 42; printf(\"%d\\n\", *p); free(p); return 0;", expect: ["42"] },
    calloc_memory_starts_zeroed_for_int_slots => { body: "int *p = (int *)calloc(2, sizeof(int)); printf(\"%d %d\\n\", p[0], p[1]); free(p); return 0;", expect: ["0 0"] },
    free_can_follow_malloc_after_write => { body: "char *p = (char *)malloc(4); p[0] = 'o'; p[1] = 'k'; p[2] = '\\0'; puts(p); free(p); return 0;", expect: ["ok"] },
    malloc_sizeof_array_pattern_can_store_multiple_ints => { body: "int *p = (int *)malloc(3 * sizeof(int)); p[0] = 1; p[1] = 2; p[2] = 3; printf(\"%d %d %d\\n\", p[0], p[1], p[2]); free(p); return 0;", expect: ["1 2 3"] },
    calloc_char_buffer_can_hold_string_then_free => { body: "char *p = (char *)calloc(8, sizeof(char)); p[0] = 'h'; p[1] = 'i'; puts(p); free(p); return 0;", expect: ["hi"] },
    atoi_result_can_feed_arithmetic => { body: "printf(\"%d\\n\", atoi(\"40\") + 2);\nreturn 0;", expect: ["42"] },
    atol_result_can_feed_comparison => { body: "printf(\"%d\\n\", atol(\"99\") > 50);\nreturn 0;", expect: ["1"] },
    atof_result_can_feed_multiplication => { body: "printf(\"%.1f\\n\", atof(\"2.0\") * 4);\nreturn 0;", expect: ["8.0"] },
    abs_can_normalize_negative_expression => { body: "printf(\"%d\\n\", abs(3 - 8));\nreturn 0;", expect: ["5"] },
    labs_can_normalize_negative_expression => { body: "printf(\"%d\\n\", labs(3 - 8));\nreturn 0;", expect: ["5"] },
    malloc_zero_bytes_is_safely_checkable => { body: "void *p = malloc(0); printf(\"%d\\n\", p == NULL || p != NULL); free(p); return 0;", expect: ["1"] },
    calloc_zero_elements_is_safely_checkable => { body: "void *p = calloc(0, sizeof(int)); printf(\"%d\\n\", p == NULL || p != NULL); free(p); return 0;", expect: ["1"] },
    exit_can_follow_previous_output => { body: "puts(\"before\"); exit(0); puts(\"after\"); return 0;", expect: ["before"] },
    atoi_handles_embedded_space_after_digits => { body: "printf(\"%d\\n\", atoi(\"17 99\"));\nreturn 0;", expect: ["17"] },
    atof_handles_negative_fractional_text => { body: "printf(\"%.2f\\n\", atof(\"-0.75\"));\nreturn 0;", expect: ["-0.75"] },
    rand_max_can_be_printed => { body: "printf(\"%d\\n\", RAND_MAX);\nreturn 0;", expect: ["2147483647"] },
    null_macro_can_initialize_pointer => { body: "int *p = NULL; printf(\"%d\\n\", p == 0); return 0;", expect: ["1"] },
    malloc_pointer_can_be_cast_to_char_pointer_and_indexed => { body: "char *p = (char *)malloc(3); p[0] = 'a'; p[1] = 'b'; p[2] = '\\0'; puts(p); free(p); return 0;", expect: ["ab"] },
    calloc_pointer_can_be_cast_to_int_pointer_and_summed => { body: "int *p = (int *)calloc(3, sizeof(int)); p[1] = 4; p[2] = 5; printf(\"%d\\n\", p[0] + p[1] + p[2]); free(p); return 0;", expect: ["9"] },
    atoi_empty_string_yields_zero => { body: "printf(\"%d\\n\", atoi(\"\"));\nreturn 0;", expect: ["0"] },
    atol_empty_string_yields_zero => { body: "printf(\"%d\\n\", atol(\"\"));\nreturn 0;", expect: ["0"] },
    atof_empty_string_yields_zero => { body: "printf(\"%.1f\\n\", atof(\"\"));\nreturn 0;", expect: ["0.0"] },
    free_null_pointer_is_safe_to_call => { body: "free(NULL); puts(\"ok\"); return 0;", expect: ["ok"] },
    atoi_can_parse_zero => { body: "printf(\"%d\\n\", atoi(\"0\"));\nreturn 0;", expect: ["0"] },
    atof_can_parse_integer_text_as_double => { body: "printf(\"%.1f\\n\", atof(\"5\"));\nreturn 0;", expect: ["5.0"] },
    malloc_and_free_can_be_used_in_loop => { body: "for (int i = 0; i < 2; i++) { int *p = (int *)malloc(sizeof(int)); *p = i; printf(\"%d\\n\", *p); free(p); } return 0;", expect: ["0", "1"] },
    calloc_can_allocate_space_for_doubles => { body: "double *p = (double *)calloc(2, sizeof(double)); p[1] = 2.5; printf(\"%.1f\\n\", p[1]); free(p); return 0;", expect: ["2.5"] }
}
