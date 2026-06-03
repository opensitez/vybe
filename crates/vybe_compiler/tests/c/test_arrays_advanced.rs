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
    partial_integer_array_initialization_zero_fills_tail => { declarations: "int arr[4] = {1, 2};", body: "printf(\"%d %d %d %d\\n\", arr[0], arr[1], arr[2], arr[3]);\nreturn 0;", expect: ["1 2 0 0"] },
    char_array_from_string_literal_exposes_characters => { declarations: "char text[] = \"cat\";", body: "printf(\"%c %c %c\\n\", text[0], text[1], text[2]);\nreturn 0;", expect: ["c a t"] },
    char_array_from_string_literal_has_null_terminator => { declarations: "char text[] = \"cat\";", body: "printf(\"%d\\n\", text[3]);\nreturn 0;", expect: ["0"] },
    multidimensional_array_preserves_row_major_order => { declarations: "int m[2][2] = {{1, 2}, {3, 4}};", body: "printf(\"%d %d %d %d\\n\", m[0][0], m[0][1], m[1][0], m[1][1]);\nreturn 0;", expect: ["1 2 3 4"] },
    array_element_assignment_overwrites_slot => { declarations: "int arr[3] = {1, 2, 3};", body: "arr[1] = 9;\nprintf(\"%d\\n\", arr[1]);\nreturn 0;", expect: ["9"] },
    array_element_expression_reads_updated_slot => { declarations: "int arr[3] = {1, 2, 3};", body: "arr[1] = arr[0] + arr[2];\nprintf(\"%d\\n\", arr[1]);\nreturn 0;", expect: ["4"] },
    array_name_decays_to_pointer_in_dereference => { declarations: "int arr[3] = {4, 5, 6};", body: "printf(\"%d\\n\", *arr);\nreturn 0;", expect: ["4"] },
    array_indexing_matches_pointer_arithmetic => { declarations: "int arr[3] = {4, 5, 6};", body: "printf(\"%d\\n\", *(arr + 1));\nreturn 0;", expect: ["5"] },
    pointer_style_indexing_from_offset_pointer_reads_element => { declarations: "int arr[4] = {7, 8, 9, 10}; int *p = arr + 2;", body: "printf(\"%d\\n\", p[1]);\nreturn 0;", expect: ["10"] },
    sizeof_array_counts_entire_storage => { declarations: "int arr[4] = {0, 0, 0, 0};", body: "printf(\"%d\\n\", (int)sizeof(arr));\nreturn 0;", expect: ["16"] },
    sizeof_array_element_counts_single_slot => { declarations: "int arr[4] = {0, 0, 0, 0};", body: "printf(\"%d\\n\", (int)sizeof(arr[0]));\nreturn 0;", expect: ["4"] },
    array_of_chars_can_be_printed_with_puts => { declarations: "char text[] = \"vybe\";", body: "puts(text);\nreturn 0;", expect: ["vybe"] },
    multidimensional_array_can_be_summed => { declarations: "int m[2][2] = {{1, 2}, {3, 4}};", body: "printf(\"%d\\n\", m[0][0] + m[0][1] + m[1][0] + m[1][1]);\nreturn 0;", expect: ["10"] },
    array_initializer_can_use_expressions => { declarations: "int arr[3] = {1 + 1, 2 * 3, 9 / 3};", body: "printf(\"%d %d %d\\n\", arr[0], arr[1], arr[2]);\nreturn 0;", expect: ["2 6 3"] },
    array_value_can_drive_for_loop_total => { declarations: "int arr[4] = {1, 3, 5, 7};", body: "int total = 0;\nfor (int i = 0; i < 4; i++) total += arr[i];\nprintf(\"%d\\n\", total);\nreturn 0;", expect: ["16"] },
    array_reverse_index_reads_last_element => { declarations: "int arr[4] = {1, 3, 5, 7};", body: "printf(\"%d\\n\", arr[3 - 1]);\nreturn 0;", expect: ["5"] },
    char_array_element_can_be_reassigned => { declarations: "char text[] = \"cat\";", body: "text[1] = 'o';\nputs(text);\nreturn 0;", expect: ["cot"] },
    integer_array_can_store_negative_values => { declarations: "int arr[3] = {-1, -2, -3};", body: "printf(\"%d\\n\", arr[2]);\nreturn 0;", expect: ["-3"] },
    array_of_doubles_keeps_fractional_values => { declarations: "double arr[3] = {1.5, 2.5, 3.5};", body: "printf(\"%.1f\\n\", arr[1]);\nreturn 0;", expect: ["2.5"] },
    array_assignment_by_index_can_copy_other_index => { declarations: "int arr[3] = {4, 8, 0};", body: "arr[2] = arr[1];\nprintf(\"%d\\n\", arr[2]);\nreturn 0;", expect: ["8"] },
    multidimensional_array_index_expression_can_use_variable => { declarations: "int m[2][2] = {{2, 4}, {6, 8}}; int row = 1;", body: "printf(\"%d\\n\", m[row][0]);\nreturn 0;", expect: ["6"] },
    string_literal_array_size_includes_null => { declarations: "char text[] = \"cat\";", body: "printf(\"%d\\n\", (int)sizeof(text));\nreturn 0;", expect: ["4"] },
    array_name_in_boolean_context_is_truthy => { declarations: "int arr[1] = {5};", body: "if (arr) puts(\"true\"); else puts(\"false\");\nreturn 0;", expect: ["true"] },
    two_dimensional_array_row_can_decay_to_pointer => { declarations: "int m[2][2] = {{2, 4}, {6, 8}};", body: "int *row = m[1];\nprintf(\"%d %d\\n\", row[0], row[1]);\nreturn 0;", expect: ["6 8"] },
    array_initializer_can_mix_hex_and_decimal => { declarations: "int arr[3] = {0x10, 10, 010};", body: "printf(\"%d %d %d\\n\", arr[0], arr[1], arr[2]);\nreturn 0;", expect: ["16 10 8"] },
    array_copy_via_loop_moves_all_values => { declarations: "int src[3] = {3, 6, 9}; int dst[3] = {0, 0, 0};", body: "for (int i = 0; i < 3; i++) dst[i] = src[i];\nprintf(\"%d %d %d\\n\", dst[0], dst[1], dst[2]);\nreturn 0;", expect: ["3 6 9"] },
    char_array_can_be_iterated_until_null_terminator => { declarations: "char text[] = \"go\";", body: "int i = 0;\nwhile (text[i]) { printf(\"%c\\n\", text[i]); i++; }\nreturn 0;", expect: ["g", "o"] },
    array_element_postfix_increment_updates_storage => { declarations: "int arr[2] = {2, 4};", body: "arr[0]++;\nprintf(\"%d\\n\", arr[0]);\nreturn 0;", expect: ["3"] },
    array_can_hold_enum_like_constants => { declarations: "int arr[3] = {100, 200, 300};", body: "printf(\"%d\\n\", arr[1]);\nreturn 0;", expect: ["200"] },
    pointer_to_first_array_element_can_be_indexed => { declarations: "int arr[3] = {9, 8, 7}; int *p = &arr[0];", body: "printf(\"%d\\n\", p[2]);\nreturn 0;", expect: ["7"] }
}