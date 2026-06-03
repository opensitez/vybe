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
    forward_declaration_allows_call_before_definition => { declarations: "int add(int a, int b);\nint add(int a, int b) { return a + b; }", body: "printf(\"%d\\n\", add(3, 4));\nreturn 0;", expect: ["7"] },
    function_can_return_double_value => { declarations: "double half(double x) { return x / 2.0; }", body: "printf(\"%.2f\\n\", half(9.0));\nreturn 0;", expect: ["4.50"] },
    function_can_take_array_parameter => { declarations: "int first(int arr[]) { return arr[0]; }", body: "int values[3] = {8, 9, 10};\nprintf(\"%d\\n\", first(values));\nreturn 0;", expect: ["8"] },
    function_can_take_pointer_parameter_and_mutate_caller => { declarations: "void set_to_ten(int *p) { *p = 10; }", body: "int value = 3;\nset_to_ten(&value);\nprintf(\"%d\\n\", value);\nreturn 0;", expect: ["10"] },
    function_can_return_pointer_to_array_element => { declarations: "int *second(int arr[]) { return &arr[1]; }", body: "int values[3] = {8, 9, 10};\nprintf(\"%d\\n\", *second(values));\nreturn 0;", expect: ["9"] },
    function_can_call_another_function => { declarations: "int square(int x) { return x * x; }\nint twice_square(int x) { return square(x) * 2; }", body: "printf(\"%d\\n\", twice_square(3));\nreturn 0;", expect: ["18"] },
    function_can_use_local_variables_without_leaking_state => { declarations: "int next_two_sum(int x) { int y = x + 1; return x + y; }", body: "printf(\"%d\\n\", next_two_sum(4));\nreturn 0;", expect: ["9"] },
    function_can_return_early_from_branch => { declarations: "int sign_label(int x) { if (x < 0) return -1; if (x > 0) return 1; return 0; }", body: "printf(\"%d %d %d\\n\", sign_label(-4), sign_label(0), sign_label(4));\nreturn 0;", expect: ["-1 0 1"] },
    recursion_can_handle_countdown_sum => { declarations: "int sum_to(int n) { if (n <= 0) return 0; return n + sum_to(n - 1); }", body: "printf(\"%d\\n\", sum_to(5));\nreturn 0;", expect: ["15"] },
    function_can_return_struct_like_multiple_outputs_via_pointers => { declarations: "void divmod(int a, int b, int *q, int *r) { *q = a / b; *r = a % b; }", body: "int q = 0; int r = 0;\ndivmod(17, 5, &q, &r);\nprintf(\"%d %d\\n\", q, r);\nreturn 0;", expect: ["3 2"] },
    function_with_void_parameter_list_can_be_called => { declarations: "int answer(void) { return 42; }", body: "printf(\"%d\\n\", answer());\nreturn 0;", expect: ["42"] },
    function_can_accept_char_pointer_and_print_string => { declarations: "void greet(char *name) { printf(\"hi %s\\n\", name); }", body: "greet(\"vybe\");\nreturn 0;", expect: ["hi vybe"] },
    function_can_accept_double_and_int_mix => { declarations: "double scale(double x, int factor) { return x * factor; }", body: "printf(\"%.1f\\n\", scale(2.5, 3));\nreturn 0;", expect: ["7.5"] },
    function_parameter_shadowing_does_not_change_global_named_variable => { declarations: "int value = 5;\nint identity(int value) { return value; }", body: "printf(\"%d %d\\n\", identity(9), value);\nreturn 0;", expect: ["9 5"] },
    function_can_use_static_local_to_persist_state => { declarations: "int next_counter(void) { static int value = 0; value++; return value; }", body: "printf(\"%d\\n\", next_counter());\nprintf(\"%d\\n\", next_counter());\nreturn 0;", expect: ["1", "2"] },
    function_can_return_char_from_integer_code => { declarations: "char next_letter(char c) { return c + 1; }", body: "printf(\"%c\\n\", next_letter('a'));\nreturn 0;", expect: ["b"] },
    function_can_receive_array_and_sum_with_length => { declarations: "int sum(int arr[], int len) { int total = 0; for (int i = 0; i < len; i++) total += arr[i]; return total; }", body: "int data[4] = {1, 2, 3, 4};\nprintf(\"%d\\n\", sum(data, 4));\nreturn 0;", expect: ["10"] },
    function_can_return_pointer_argument_for_chaining => { declarations: "int *write_value(int *p, int value) { *p = value; return p; }", body: "int x = 1;\nprintf(\"%d\\n\", *write_value(&x, 9));\nreturn 0;", expect: ["9"] },
    function_can_take_function_result_as_argument => { declarations: "int add_one(int x) { return x + 1; }\nint double_it(int x) { return x * 2; }", body: "printf(\"%d\\n\", double_it(add_one(4)));\nreturn 0;", expect: ["10"] },
    function_can_have_multiple_return_paths_with_same_type => { declarations: "double clamp_unit(double x) { if (x < 0.0) return 0.0; if (x > 1.0) return 1.0; return x; }", body: "printf(\"%.1f %.1f %.1f\\n\", clamp_unit(-1.0), clamp_unit(0.5), clamp_unit(2.0));\nreturn 0;", expect: ["0.0 0.5 1.0"] },
    mutual_calls_can_be_ordered_with_prototypes => { declarations: "int odd(int n);\nint even(int n) { return n == 0 ? 1 : odd(n - 1); }\nint odd(int n) { return n == 0 ? 0 : even(n - 1); }", body: "printf(\"%d %d\\n\", even(4), odd(4));\nreturn 0;", expect: ["1 0"] },
    function_can_take_pointer_to_char_and_return_offset => { declarations: "char *tail(char *text) { return text + 2; }", body: "puts(tail(\"hello\"));\nreturn 0;", expect: ["llo"] },
    function_can_mutate_array_element_through_parameter => { declarations: "void set_second(int arr[]) { arr[1] = 99; }", body: "int arr[3] = {1, 2, 3};\nset_second(arr);\nprintf(\"%d\\n\", arr[1]);\nreturn 0;", expect: ["99"] },
    function_can_return_result_of_comparison => { declarations: "int is_even(int x) { return x % 2 == 0; }", body: "printf(\"%d %d\\n\", is_even(4), is_even(5));\nreturn 0;", expect: ["1 0"] },
    function_can_consume_pointer_to_pointer => { declarations: "int read_twice(int **pp) { return **pp; }", body: "int x = 6; int *p = &x;\nprintf(\"%d\\n\", read_twice(&p));\nreturn 0;", expect: ["6"] },
    function_can_return_array_element_by_index_parameter => { declarations: "int get_at(int arr[], int index) { return arr[index]; }", body: "int arr[3] = {2, 4, 6};\nprintf(\"%d\\n\", get_at(arr, 2));\nreturn 0;", expect: ["6"] },
    function_can_accumulate_using_static_local_across_calls => { declarations: "int add_and_keep(int x) { static int total = 0; total += x; return total; }", body: "printf(\"%d\\n\", add_and_keep(3));\nprintf(\"%d\\n\", add_and_keep(4));\nreturn 0;", expect: ["3", "7"] },
    function_can_return_double_from_integer_arguments => { declarations: "double average(int a, int b) { return (a + b) / 2.0; }", body: "printf(\"%.1f\\n\", average(3, 4));\nreturn 0;", expect: ["3.5"] },
    function_can_read_global_variable => { declarations: "int base = 10;\nint plus_base(int x) { return base + x; }", body: "printf(\"%d\\n\", plus_base(5));\nreturn 0;", expect: ["15"] },
    function_can_shadow_global_without_mutating_it => { declarations: "int x = 5;\nint local_shadow(void) { int x = 8; return x; }", body: "printf(\"%d %d\\n\", local_shadow(), x);\nreturn 0;", expect: ["8 5"] }
}