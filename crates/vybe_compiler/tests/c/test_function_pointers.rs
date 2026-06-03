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
    function_pointer_variable_can_call_target => { declarations: "int add_one(int x) { return x + 1; }", body: "int (*fp)(int) = add_one;\nprintf(\"%d\\n\", fp(4));\nreturn 0;", expect: ["5"] },
    function_pointer_can_be_assigned_by_address => { declarations: "int add_one(int x) { return x + 1; }", body: "int (*fp)(int) = &add_one;\nprintf(\"%d\\n\", fp(4));\nreturn 0;", expect: ["5"] },
    function_pointer_can_be_passed_as_argument => { declarations: "int add_one(int x) { return x + 1; }\nint apply(int (*fp)(int), int value) { return fp(value); }", body: "printf(\"%d\\n\", apply(add_one, 7));\nreturn 0;", expect: ["8"] },
    function_pointer_can_select_between_two_functions => { declarations: "int add_one(int x) { return x + 1; }\nint double_it(int x) { return x * 2; }", body: "int (*fp)(int) = double_it;\nprintf(\"%d\\n\", fp(7));\nreturn 0;", expect: ["14"] },
    function_pointer_can_be_stored_in_array => { declarations: "int add_one(int x) { return x + 1; }\nint double_it(int x) { return x * 2; }", body: "int (*ops[2])(int) = {add_one, double_it};\nprintf(\"%d %d\\n\", ops[0](3), ops[1](3));\nreturn 0;", expect: ["4 6"] },
    function_pointer_can_reference_void_function => { declarations: "void greet(void) { puts(\"hi\"); }", body: "void (*fp)(void) = greet;\nfp();\nreturn 0;", expect: ["hi"] },
    function_pointer_parameter_can_be_invoked_twice => { declarations: "int add_one(int x) { return x + 1; }\nint apply_twice(int (*fp)(int), int value) { return fp(fp(value)); }", body: "printf(\"%d\\n\", apply_twice(add_one, 3));\nreturn 0;", expect: ["5"] },
    typedef_like_function_pointer_usage_without_typedef_still_calls => { declarations: "int square(int x) { return x * x; }", body: "int (*fp)(int) = square;\nprintf(\"%d\\n\", (*fp)(5));\nreturn 0;", expect: ["25"] },
    function_pointer_can_point_to_function_returning_double => { declarations: "double half(double x) { return x / 2.0; }", body: "double (*fp)(double) = half;\nprintf(\"%.2f\\n\", fp(9.0));\nreturn 0;", expect: ["4.50"] },
    function_pointer_can_switch_targets_dynamically => { declarations: "int add_one(int x) { return x + 1; }\nint sub_one(int x) { return x - 1; }", body: "int (*fp)(int) = add_one;\nprintf(\"%d\\n\", fp(10));\nfp = sub_one;\nprintf(\"%d\\n\", fp(10));\nreturn 0;", expect: ["11", "9"] },
    function_pointer_can_be_compared_to_function_symbol => { declarations: "int add_one(int x) { return x + 1; }", body: "int (*fp)(int) = add_one;\nprintf(\"%d\\n\", fp == add_one);\nreturn 0;", expect: ["1"] },
    function_pointer_can_be_returned_from_function => { declarations: "int add_one(int x) { return x + 1; }\nint (*pick(void))(int) { return add_one; }", body: "printf(\"%d\\n\", pick()(4));\nreturn 0;", expect: ["5"] },
    function_pointer_can_be_nested_in_struct => { declarations: "struct Op { int (*apply)(int); };\nint add_one(int x) { return x + 1; }", body: "struct Op op = {add_one};\nprintf(\"%d\\n\", op.apply(4));\nreturn 0;", expect: ["5"] },
    function_pointer_array_index_can_select_behavior => { declarations: "int add_one(int x) { return x + 1; }\nint double_it(int x) { return x * 2; }", body: "int (*ops[2])(int) = {add_one, double_it}; int index = 1;\nprintf(\"%d\\n\", ops[index](5));\nreturn 0;", expect: ["10"] },
    function_pointer_parameter_can_use_local_choice => { declarations: "int add_one(int x) { return x + 1; }\nint double_it(int x) { return x * 2; }\nint call_selected(int which, int value) { int (*fp)(int) = which ? double_it : add_one; return fp(value); }", body: "printf(\"%d %d\\n\", call_selected(0, 3), call_selected(1, 3));\nreturn 0;", expect: ["4 6"] },
    function_pointer_to_char_function_can_return_character => { declarations: "char next_letter(char c) { return c + 1; }", body: "char (*fp)(char) = next_letter;\nprintf(\"%c\\n\", fp('a'));\nreturn 0;", expect: ["b"] },
    function_pointer_to_predicate_can_drive_if => { declarations: "int is_even(int x) { return x % 2 == 0; }", body: "int (*pred)(int) = is_even;\nif (pred(6)) puts(\"even\"); else puts(\"odd\");\nreturn 0;", expect: ["even"] },
    function_pointer_can_be_dereferenced_explicitly_before_call => { declarations: "int add_one(int x) { return x + 1; }", body: "int (*fp)(int) = add_one;\nprintf(\"%d\\n\", (*fp)(9));\nreturn 0;", expect: ["10"] },
    function_pointer_and_function_symbol_can_share_same_behavior => { declarations: "int triple(int x) { return x * 3; }", body: "int (*fp)(int) = triple;\nprintf(\"%d %d\\n\", triple(3), fp(3));\nreturn 0;", expect: ["9 9"] },
    function_pointer_can_be_null_checked_before_call => { declarations: "int (*fp)(int) = NULL;", body: "if (!fp) puts(\"null\"); else puts(\"bad\");\nreturn 0;", expect: ["null"] },
    function_pointer_can_accept_pointer_argument_type => { declarations: "int read_ptr(int *p) { return *p; }", body: "int x = 12; int (*fp)(int *) = read_ptr;\nprintf(\"%d\\n\", fp(&x));\nreturn 0;", expect: ["12"] },
    function_pointer_can_target_function_with_two_parameters => { declarations: "int add(int a, int b) { return a + b; }", body: "int (*fp)(int, int) = add;\nprintf(\"%d\\n\", fp(4, 5));\nreturn 0;", expect: ["9"] },
    function_pointer_can_live_in_array_of_length_one => { declarations: "int negate(int x) { return -x; }", body: "int (*ops[1])(int) = {negate};\nprintf(\"%d\\n\", ops[0](5));\nreturn 0;", expect: ["-5"] },
    function_pointer_returning_pointer_can_be_called => { declarations: "int *identity(int *p) { return p; }", body: "int x = 7; int *(*fp)(int *) = identity;\nprintf(\"%d\\n\", *fp(&x));\nreturn 0;", expect: ["7"] },
    function_pointer_can_call_recursive_target => { declarations: "int fact(int n) { return n <= 1 ? 1 : n * fact(n - 1); }", body: "int (*fp)(int) = fact;\nprintf(\"%d\\n\", fp(5));\nreturn 0;", expect: ["120"] }
}