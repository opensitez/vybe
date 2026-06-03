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
    pointer_plus_one_moves_to_next_int_element => { declarations: "int arr[3] = {4, 5, 6}; int *p = arr;", body: "printf(\"%d\\n\", *(p + 1));\nreturn 0;", expect: ["5"] },
    pointer_minus_one_moves_back_to_previous_int_element => { declarations: "int arr[3] = {4, 5, 6}; int *p = &arr[2];", body: "printf(\"%d\\n\", *(p - 1));\nreturn 0;", expect: ["5"] },
    pointer_increment_advances_to_next_element => { declarations: "int arr[3] = {4, 5, 6}; int *p = arr;", body: "p++;\nprintf(\"%d\\n\", *p);\nreturn 0;", expect: ["5"] },
    pointer_decrement_rewinds_to_previous_element => { declarations: "int arr[3] = {4, 5, 6}; int *p = &arr[2];", body: "p--;\nprintf(\"%d\\n\", *p);\nreturn 0;", expect: ["5"] },
    pointer_difference_counts_elements_not_bytes => { declarations: "int arr[5] = {0, 1, 2, 3, 4};", body: "printf(\"%d\\n\", (int)(&arr[4] - &arr[1]));\nreturn 0;", expect: ["3"] },
    pointer_addition_can_skip_multiple_elements => { declarations: "int arr[5] = {0, 1, 2, 3, 4}; int *p = arr;", body: "p = p + 3;\nprintf(\"%d\\n\", *p);\nreturn 0;", expect: ["3"] },
    pointer_subtraction_can_skip_multiple_elements => { declarations: "int arr[5] = {0, 1, 2, 3, 4}; int *p = &arr[4];", body: "p = p - 2;\nprintf(\"%d\\n\", *p);\nreturn 0;", expect: ["2"] },
    char_pointer_arithmetic_moves_by_single_byte_character => { declarations: "char text[] = \"hello\"; char *p = text;", body: "p += 2;\nprintf(\"%c\\n\", *p);\nreturn 0;", expect: ["l"] },
    pointer_subscript_matches_arithmetic_form => { declarations: "int arr[4] = {2, 4, 6, 8}; int *p = arr;", body: "printf(\"%d %d\\n\", p[3], *(p + 3));\nreturn 0;", expect: ["8 8"] },
    reverse_subscript_form_is_equivalent_in_c => { declarations: "int arr[4] = {2, 4, 6, 8}; int *p = arr;", body: "printf(\"%d\\n\", 2[p]);\nreturn 0;", expect: ["6"] },
    pointer_walk_can_sum_array_values => { declarations: "int arr[4] = {1, 2, 3, 4}; int *p = arr;", body: "int total = 0;\nfor (int i = 0; i < 4; i++) total += *(p + i);\nprintf(\"%d\\n\", total);\nreturn 0;", expect: ["10"] },
    postfix_pointer_increment_uses_old_location_then_advances => { declarations: "int arr[3] = {7, 8, 9}; int *p = arr;", body: "printf(\"%d\\n\", *p++);\nprintf(\"%d\\n\", *p);\nreturn 0;", expect: ["7", "8"] },
    prefix_pointer_increment_advances_before_dereference => { declarations: "int arr[3] = {7, 8, 9}; int *p = arr;", body: "printf(\"%d\\n\", *++p);\nreturn 0;", expect: ["8"] },
    postfix_pointer_decrement_uses_old_location_then_rewinds => { declarations: "int arr[3] = {7, 8, 9}; int *p = &arr[2];", body: "printf(\"%d\\n\", *p--);\nprintf(\"%d\\n\", *p);\nreturn 0;", expect: ["9", "8"] },
    prefix_pointer_decrement_rewinds_before_dereference => { declarations: "int arr[3] = {7, 8, 9}; int *p = &arr[2];", body: "printf(\"%d\\n\", *--p);\nreturn 0;", expect: ["8"] },
    pointer_difference_on_char_array_counts_characters => { declarations: "char text[] = \"hello\";", body: "printf(\"%d\\n\", (int)(&text[4] - &text[1]));\nreturn 0;", expect: ["3"] },
    pointer_loop_can_emit_each_character => { declarations: "char text[] = \"go\"; char *p = text;", body: "while (*p) { printf(\"%c\\n\", *p); p++; }\nreturn 0;", expect: ["g", "o"] },
    pointer_subtraction_between_same_array_bounds_can_measure_length => { declarations: "int arr[4] = {1, 2, 3, 4};", body: "printf(\"%d\\n\", (int)(&arr[4] - &arr[0]));\nreturn 0;", expect: ["4"] },
    pointer_plus_zero_leaves_address_unchanged => { declarations: "int arr[2] = {5, 6}; int *p = arr;", body: "printf(\"%d\\n\", *(p + 0));\nreturn 0;", expect: ["5"] },
    pointer_arithmetic_can_target_last_character_before_null => { declarations: "char text[] = \"cat\"; char *p = text + 2;", body: "printf(\"%c\\n\", *p);\nreturn 0;", expect: ["t"] },
    pointer_arithmetic_can_access_null_terminator => { declarations: "char text[] = \"cat\"; char *p = text + 3;", body: "printf(\"%d\\n\", *p);\nreturn 0;", expect: ["0"] },
    subtraction_of_pointers_to_adjacent_ints_is_one => { declarations: "int arr[3] = {1, 2, 3};", body: "printf(\"%d\\n\", (int)(&arr[2] - &arr[1]));\nreturn 0;", expect: ["1"] },
    pointer_addition_inside_subscript_expression_reads_target => { declarations: "int arr[4] = {10, 20, 30, 40}; int *p = arr;", body: "printf(\"%d\\n\", (p + 2)[0]);\nreturn 0;", expect: ["30"] },
    pointer_walk_back_and_forth_returns_to_same_value => { declarations: "int arr[4] = {10, 20, 30, 40}; int *p = &arr[1];", body: "p += 2;\np -= 2;\nprintf(\"%d\\n\", *p);\nreturn 0;", expect: ["20"] },
    pointer_difference_can_drive_loop_bound => { declarations: "int arr[4] = {1, 2, 3, 4}; int *start = &arr[0]; int *end = &arr[4];", body: "printf(\"%d\\n\", (int)(end - start));\nreturn 0;", expect: ["4"] },
    pointer_increment_on_char_pointer_can_print_suffix => { declarations: "char text[] = \"world\"; char *p = text;", body: "p++;\nputs(p);\nreturn 0;", expect: ["orld"] },
    pointer_decrement_from_one_past_end_reaches_last_element => { declarations: "int arr[3] = {4, 5, 6}; int *p = &arr[3];", body: "p--;\nprintf(\"%d\\n\", *p);\nreturn 0;", expect: ["6"] },
    pointer_indexing_after_increment_reads_new_position => { declarations: "int arr[4] = {4, 5, 6, 7}; int *p = arr;", body: "p++;\nprintf(\"%d\\n\", p[1]);\nreturn 0;", expect: ["6"] },
    pointer_to_middle_element_supports_negative_offset_via_subtraction => { declarations: "int arr[4] = {4, 5, 6, 7}; int *p = &arr[2];", body: "printf(\"%d\\n\", *(p - 2));\nreturn 0;", expect: ["4"] },
    pointer_arithmetic_with_double_array_advances_by_element => { declarations: "double arr[3] = {1.5, 2.5, 3.5}; double *p = arr;", body: "printf(\"%.1f\\n\", *(p + 2));\nreturn 0;", expect: ["3.5"] }
}