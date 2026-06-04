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
    address_of_and_dereference_round_trip_value => { declarations: "int x = 7; int *p = &x;", body: "printf(\"%d\\n\", *p);\nreturn 0;", expect: ["7"] },
    dereference_can_update_original_variable => { declarations: "int x = 7; int *p = &x;", body: "*p = 11;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["11"] },
    pointer_aliases_same_storage => { declarations: "int x = 5; int *a = &x; int *b = &x;", body: "*a = 9;\nprintf(\"%d\\n\", *b);\nreturn 0;", expect: ["9"] },
    pointer_can_point_to_array_first_element => { declarations: "int arr[3] = {4, 5, 6}; int *p = arr;", body: "printf(\"%d\\n\", *p);\nreturn 0;", expect: ["4"] },
    pointer_can_point_to_character_array => { declarations: "char text[] = \"vybe\"; char *p = text;", body: "printf(\"%c\\n\", *p);\nreturn 0;", expect: ["v"] },
    null_pointer_compares_equal_to_zero_constant => { declarations: "int *p = 0;", body: "if (p == NULL) puts(\"null\"); else puts(\"bad\");\nreturn 0;", expect: ["null"] },
    pointer_to_pointer_can_follow_two_levels => { declarations: "int x = 7; int *p = &x; int **pp = &p;", body: "printf(\"%d\\n\", **pp);\nreturn 0;", expect: ["7"] },
    pointer_to_pointer_can_update_original_value => { declarations: "int x = 7; int *p = &x; int **pp = &p;", body: "**pp = 12;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["12"] },
    dereference_of_address_expression_is_identity => { declarations: "int x = 42;", body: "printf(\"%d\\n\", *&x);\nreturn 0;", expect: ["42"] },
    pointer_can_be_reassigned_to_other_variable => { declarations: "int a = 1; int b = 2; int *p = &a;", body: "p = &b;\nprintf(\"%d\\n\", *p);\nreturn 0;", expect: ["2"] },
    pointer_value_can_drive_if_condition => { declarations: "int x = 3; int *p = &x;", body: "if (p) puts(\"true\"); else puts(\"false\");\nreturn 0;", expect: ["true"] },
    pointer_to_char_can_be_passed_to_puts => { declarations: "char text[] = \"hello\"; char *p = text;", body: "puts(p);\nreturn 0;", expect: ["hello"] },
    multiple_dereference_reads_array_value => { declarations: "int arr[2] = {8, 9}; int *p = arr;", body: "printf(\"%d %d\\n\", *p, *(p + 1));\nreturn 0;", expect: ["8 9"] },
    pointer_comparison_same_address_is_true => { declarations: "int x = 3; int *a = &x; int *b = &x;", body: "printf(\"%d\\n\", a == b);\nreturn 0;", expect: ["1"] },
    pointer_comparison_different_addresses_is_false => { declarations: "int a = 3; int b = 4; int *pa = &a; int *pb = &b;", body: "printf(\"%d\\n\", pa == pb);\nreturn 0;", expect: ["0"] },
    pointer_difference_of_same_element_is_zero => { declarations: "int arr[3] = {1, 2, 3};", body: "printf(\"%d\\n\", (int)(&arr[1] - &arr[1]));\nreturn 0;", expect: ["0"] },
    pointer_can_reference_double_value => { declarations: "double x = 2.5; double *p = &x;", body: "printf(\"%.1f\\n\", *p);\nreturn 0;", expect: ["2.5"] },
    pointer_can_reference_char_value => { declarations: "char c = 'q'; char *p = &c;", body: "printf(\"%c\\n\", *p);\nreturn 0;", expect: ["q"] },
    pointer_write_through_char_updates_character => { declarations: "char c = 'q'; char *p = &c;", body: "*p = 'r';\nprintf(\"%c\\n\", c);\nreturn 0;", expect: ["r"] },
    pointer_to_struct_member_address_reads_field => { declarations: "struct Pair { int a; int b; }; struct Pair pair = {2, 4}; int *p = &pair.b;", body: "printf(\"%d\\n\", *p);\nreturn 0;", expect: ["4"] },
    pointer_can_be_initialized_from_array_subscript_address => { declarations: "int arr[3] = {7, 8, 9}; int *p = &arr[2];", body: "printf(\"%d\\n\", *p);\nreturn 0;", expect: ["9"] },
    null_pointer_can_be_checked_before_dereference => { declarations: "int *p = NULL;", body: "if (!p) puts(\"safe\"); else puts(\"bad\");\nreturn 0;", expect: ["safe"] },
    pointer_assignment_from_other_pointer_copies_address => { declarations: "int x = 10; int *a = &x; int *b = NULL;", body: "b = a;\nprintf(\"%d\\n\", *b);\nreturn 0;", expect: ["10"] },
    pointer_to_pointer_can_be_compared_with_address_of_pointer => { declarations: "int x = 10; int *p = &x; int **pp = &p;", body: "printf(\"%d\\n\", pp == &p);\nreturn 0;", expect: ["1"] },
    pointer_indirection_on_array_slot_can_update_element => { declarations: "int arr[2] = {1, 2}; int *p = &arr[1];", body: "*p = 9;\nprintf(\"%d\\n\", arr[1]);\nreturn 0;", expect: ["9"] },
    pointer_can_follow_string_suffix => { declarations: "char text[] = \"hello\"; char *p = text + 2;", body: "puts(p);\nreturn 0;", expect: ["llo"] },
    pointer_to_first_char_can_measure_offset_with_subtraction => { declarations: "char text[] = \"hello\"; char *p = text + 3;", body: "printf(\"%d\\n\", (int)(p - text));\nreturn 0;", expect: ["3"] },
    pointer_to_array_first_element_is_same_as_array_name => { declarations: "int arr[2] = {5, 6};", body: "printf(\"%d\\n\", &arr[0] == arr);\nreturn 0;", expect: ["1"] },
    pointer_to_pointer_can_follow_reassigned_pointer => { declarations: "int a = 1; int b = 2; int *p = &a; int **pp = &p;", body: "p = &b;\nprintf(\"%d\\n\", **pp);\nreturn 0;", expect: ["2"] },
    dereference_of_parenthesized_pointer_expression_reads_value => { declarations: "int x = 14; int *p = &x;", body: "printf(\"%d\\n\", *(p));\nreturn 0;", expect: ["14"] },
    pointer_can_read_through_const_text_literal_binding => { declarations: "char *p = \"vybe\";", body: "printf(\"%c\\n\", p[2]);\nreturn 0;", expect: ["b"] },
    pointer_and_integer_zero_compare_equal => { declarations: "int *p = NULL;", body: "printf(\"%d\\n\", p == 0);\nreturn 0;", expect: ["1"] },
    pointer_and_null_macro_compare_equal => { declarations: "int *p = NULL;", body: "printf(\"%d\\n\", p == NULL);\nreturn 0;", expect: ["1"] },
    pointer_can_select_array_element_via_subscript => { declarations: "int arr[3] = {2, 4, 6}; int *p = arr;", body: "printf(\"%d\\n\", p[1]);\nreturn 0;", expect: ["4"] },
    pointer_to_local_variable_keeps_latest_value => { declarations: "int x = 1; int *p = &x;", body: "x = 8;\nprintf(\"%d\\n\", *p);\nreturn 0;", expect: ["8"] }
}
