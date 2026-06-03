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
    sizeof_char_is_one => { declarations: "", body: "printf(\"%d\\n\", (int)sizeof(char)); return 0;", expect: ["1"] },
    sizeof_int_is_four => { declarations: "", body: "printf(\"%d\\n\", (int)sizeof(int)); return 0;", expect: ["4"] },
    sizeof_double_is_eight => { declarations: "", body: "printf(\"%d\\n\", (int)sizeof(double)); return 0;", expect: ["8"] },
    sizeof_char_variable_is_one => { declarations: "char c = 'a';", body: "printf(\"%d\\n\", (int)sizeof(c)); return 0;", expect: ["1"] },
    sizeof_int_variable_is_four => { declarations: "int x = 7;", body: "printf(\"%d\\n\", (int)sizeof(x)); return 0;", expect: ["4"] },
    sizeof_double_variable_is_eight => { declarations: "double x = 7.0;", body: "printf(\"%d\\n\", (int)sizeof(x)); return 0;", expect: ["8"] },
    sizeof_array_counts_total_storage => { declarations: "int arr[5] = {0};", body: "printf(\"%d\\n\", (int)sizeof(arr)); return 0;", expect: ["20"] },
    sizeof_pointer_counts_pointer_storage => { declarations: "int x = 0; int *p = &x;", body: "printf(\"%d\\n\", (int)sizeof(p)); return 0;", expect: ["8"] },
    sizeof_pointer_target_differs_from_pointer => { declarations: "int x = 0; int *p = &x;", body: "printf(\"%d %d\\n\", (int)sizeof(p), (int)sizeof(*p)); return 0;", expect: ["8 4"] },
    sizeof_char_array_from_literal_includes_null => { declarations: "char text[] = \"abc\";", body: "printf(\"%d\\n\", (int)sizeof(text)); return 0;", expect: ["4"] },
    sizeof_struct_counts_all_fields => { declarations: "struct Pair { int a; int b; };", body: "printf(\"%d\\n\", (int)sizeof(struct Pair)); return 0;", expect: ["8"] },
    sizeof_union_counts_largest_member => { declarations: "union Data { int i; char c; };", body: "printf(\"%d\\n\", (int)sizeof(union Data)); return 0;", expect: ["4"] },
    sizeof_expression_does_not_evaluate_increment => { declarations: "int x = 3;", body: "printf(\"%d %d\\n\", (int)sizeof(x++), x); return 0;", expect: ["4 3"] },
    sizeof_multidimensional_array_counts_all_elements => { declarations: "int m[2][3] = {{0}};", body: "printf(\"%d\\n\", (int)sizeof(m)); return 0;", expect: ["24"] },
    sizeof_row_of_multidimensional_array_counts_subarray => { declarations: "int m[2][3] = {{0}};", body: "printf(\"%d\\n\", (int)sizeof(m[0])); return 0;", expect: ["12"] },
    sizeof_string_literal_expression_counts_bytes_with_null => { declarations: "", body: "printf(\"%d\\n\", (int)sizeof(\"abc\")); return 0;", expect: ["4"] },
    sizeof_char_pointer_from_literal_is_pointer_size => { declarations: "char *text = \"abc\";", body: "printf(\"%d\\n\", (int)sizeof(text)); return 0;", expect: ["8"] },
    sizeof_ternary_expression_uses_common_type => { declarations: "", body: "printf(\"%d\\n\", (int)sizeof(1 ? 1 : 2.0)); return 0;", expect: ["8"] },
    sizeof_cast_expression_uses_cast_type => { declarations: "", body: "printf(\"%d\\n\", (int)sizeof((double)1)); return 0;", expect: ["8"] },
    sizeof_array_element_expression_is_element_size => { declarations: "int arr[3] = {1,2,3};", body: "printf(\"%d\\n\", (int)sizeof(arr[1])); return 0;", expect: ["4"] },
    sizeof_struct_variable_matches_struct_type => { declarations: "struct Pair { int a; int b; }; struct Pair pair = {1,2};", body: "printf(\"%d\\n\", (int)sizeof(pair)); return 0;", expect: ["8"] },
    sizeof_union_variable_matches_union_type => { declarations: "union Data { int i; char c; }; union Data data;", body: "printf(\"%d\\n\", (int)sizeof(data)); return 0;", expect: ["4"] },
    sizeof_function_parameter_array_is_pointer_size => { declarations: "int size_of_param(int arr[]) { return (int)sizeof(arr); }", body: "int arr[3] = {0}; printf(\"%d\\n\", size_of_param(arr)); return 0;", expect: ["8"] },
    sizeof_const_qualified_int_matches_int => { declarations: "const int value = 3;", body: "printf(\"%d\\n\", (int)sizeof(value)); return 0;", expect: ["4"] },
    sizeof_pointer_difference_expression_is_integer_like_size => { declarations: "int arr[3] = {0};", body: "printf(\"%d\\n\", (int)sizeof(&arr[2] - &arr[0])); return 0;", expect: ["8"] },
    sizeof_comma_expression_uses_last_operand_type => { declarations: "", body: "printf(\"%d\\n\", (int)sizeof((1, 2.0))); return 0;", expect: ["8"] },
    sizeof_function_pointer_is_pointer_size => { declarations: "int add_one(int x) { return x + 1; } int (*fp)(int) = add_one;", body: "printf(\"%d\\n\", (int)sizeof(fp)); return 0;", expect: ["8"] },
    sizeof_void_pointer_matches_other_pointer_size => { declarations: "void *p = 0;", body: "printf(\"%d\\n\", (int)sizeof(p)); return 0;", expect: ["8"] },
    sizeof_two_dimensional_row_pointer_differs_from_whole_array => { declarations: "int m[2][3] = {{0}};", body: "printf(\"%d %d\\n\", (int)sizeof(m), (int)sizeof(&m[0])); return 0;", expect: ["24 8"] },
    sizeof_char_literal_uses_int_type_in_c => { declarations: "", body: "printf(\"%d\\n\", (int)sizeof('A')); return 0;", expect: ["4"] }
}