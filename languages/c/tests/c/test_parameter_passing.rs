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
    integer_parameter_is_passed_by_value => { declarations: "void bump(int x) { x++; }", body: "int value = 3; bump(value); printf(\"%d\\n\", value); return 0;", expect: ["3"] },
    pointer_parameter_can_mutate_caller_variable => { declarations: "void bump(int *x) { (*x)++; }", body: "int value = 3; bump(&value); printf(\"%d\\n\", value); return 0;", expect: ["4"] },
    array_parameter_decays_to_pointer => { declarations: "int second(int values[]) { return values[1]; }", body: "int values[3] = {1, 2, 3}; printf(\"%d\\n\", second(values)); return 0;", expect: ["2"] },
    array_parameter_mutation_affects_caller_storage => { declarations: "void set_second(int values[]) { values[1] = 9; }", body: "int values[3] = {1, 2, 3}; set_second(values); printf(\"%d\\n\", values[1]); return 0;", expect: ["9"] },
    struct_parameter_is_copied_by_value => { declarations: "struct Pair { int a; int b; }; void change(struct Pair pair) { pair.a = 9; }", body: "struct Pair pair = {1,2}; change(pair); printf(\"%d\\n\", pair.a); return 0;", expect: ["1"] },
    struct_pointer_parameter_can_mutate_original_struct => { declarations: "struct Pair { int a; int b; }; void change(struct Pair *pair) { pair->a = 9; }", body: "struct Pair pair = {1,2}; change(&pair); printf(\"%d\\n\", pair.a); return 0;", expect: ["9"] },
    char_pointer_parameter_can_print_text => { declarations: "void show(char *text) { puts(text); }", body: "show(\"vybe\"); return 0;", expect: ["vybe"] },
    double_parameter_preserves_fractional_value => { declarations: "double half(double x) { return x / 2.0; }", body: "printf(\"%.1f\\n\", half(5.0)); return 0;", expect: ["2.5"] },
    multiple_parameters_keep_order => { declarations: "int mix(int a, int b, int c) { return a + b * c; }", body: "printf(\"%d\\n\", mix(2, 3, 4)); return 0;", expect: ["14"] },
    pointer_to_pointer_parameter_can_follow_two_levels => { declarations: "int read2(int **pp) { return **pp; }", body: "int value = 6; int *p = &value; printf(\"%d\\n\", read2(&p)); return 0;", expect: ["6"] },
    const_pointer_parameter_can_read_values => { declarations: "int first(const int *values) { return values[0]; }", body: "int values[2] = {7, 8}; printf(\"%d\\n\", first(values)); return 0;", expect: ["7"] },
    function_pointer_parameter_can_invoke_callback => { declarations: "int apply(int (*fn)(int), int value) { return fn(value); } int add_one(int x) { return x + 1; }", body: "printf(\"%d\\n\", apply(add_one, 4)); return 0;", expect: ["5"] },
    pass_by_value_can_use_expression_argument => { declarations: "int double_it(int x) { return x * 2; }", body: "printf(\"%d\\n\", double_it(3 + 1)); return 0;", expect: ["8"] },
    pass_by_pointer_can_target_array_element => { declarations: "void set_to_ten(int *x) { *x = 10; }", body: "int values[2] = {1, 2}; set_to_ten(&values[1]); printf(\"%d\\n\", values[1]); return 0;", expect: ["10"] },
    array_length_parameter_can_limit_sum => { declarations: "int sum(int values[], int len) { int total = 0; for (int i = 0; i < len; i++) total += values[i]; return total; }", body: "int values[4] = {1,2,3,4}; printf(\"%d\\n\", sum(values, 3)); return 0;", expect: ["6"] },
    pointer_parameter_can_return_same_address => { declarations: "int *identity(int *p) { return p; }", body: "int value = 11; printf(\"%d\\n\", *identity(&value)); return 0;", expect: ["11"] },
    struct_parameter_can_feed_return_value => { declarations: "struct Pair { int a; int b; }; int total(struct Pair pair) { return pair.a + pair.b; }", body: "struct Pair pair = {4,5}; printf(\"%d\\n\", total(pair)); return 0;", expect: ["9"] },
    pointer_parameter_can_be_null_checked => { declarations: "int safe_read(int *p) { return p ? *p : -1; }", body: "printf(\"%d\\n\", safe_read(NULL)); return 0;", expect: ["-1"] },
    parameter_shadowing_does_not_change_global_name => { declarations: "int value = 5; int show(int value) { return value; }", body: "printf(\"%d %d\\n\", show(8), value); return 0;", expect: ["8 5"] },
    char_parameter_can_be_promoted_and_formatted => { declarations: "int next(char c) { return c + 1; }", body: "printf(\"%d\\n\", next('A')); return 0;", expect: ["66"] },
    void_parameter_list_function_accepts_no_arguments => { declarations: "int answer(void) { return 42; }", body: "printf(\"%d\\n\", answer()); return 0;", expect: ["42"] },
    pointer_parameter_can_swap_two_values => { declarations: "void swap(int *a, int *b) { int tmp = *a; *a = *b; *b = tmp; }", body: "int a = 1; int b = 2; swap(&a, &b); printf(\"%d %d\\n\", a, b); return 0;", expect: ["2 1"] },
    array_parameter_pointer_arithmetic_reads_expected_slot => { declarations: "int third(int *values) { return *(values + 2); }", body: "int values[3] = {3, 6, 9}; printf(\"%d\\n\", third(values)); return 0;", expect: ["9"] },
    pointer_parameter_can_write_character_buffer => { declarations: "void set_first(char *text) { text[0] = 'X'; }", body: "char text[] = \"abc\"; set_first(text); puts(text); return 0;", expect: ["Xbc"] },
    double_pointer_parameter_can_retarget_pointer => { declarations: "void point_to_second(int **pp, int *target) { *pp = target; }", body: "int a = 1; int b = 2; int *p = &a; point_to_second(&p, &b); printf(\"%d\\n\", *p); return 0;", expect: ["2"] },
    function_pointer_parameter_can_run_twice => { declarations: "int apply_twice(int (*fn)(int), int value) { return fn(fn(value)); } int inc(int x) { return x + 1; }", body: "printf(\"%d\\n\", apply_twice(inc, 3)); return 0;", expect: ["5"] },
    struct_pointer_parameter_can_read_nested_field => { declarations: "struct Point { int x; int y; }; int get_y(struct Point *point) { return point->y; }", body: "struct Point point = {3, 4}; printf(\"%d\\n\", get_y(&point)); return 0;", expect: ["4"] },
    pass_by_value_of_pointer_still_points_to_same_storage => { declarations: "int read(int *p) { return *p; }", body: "int value = 13; printf(\"%d\\n\", read(&value)); return 0;", expect: ["13"] },
    array_parameter_can_use_const_length_expression => { declarations: "int sum2(int values[], int len) { int total = 0; for (int i = 0; i < len; i++) total += values[i]; return total; }", body: "enum { LEN = 2 }; int values[LEN] = {5, 6}; printf(\"%d\\n\", sum2(values, LEN)); return 0;", expect: ["11"] },
    parameter_evaluation_keeps_independent_arguments => { declarations: "int combine(int a, int b) { return a * 10 + b; }", body: "int x = 1; printf(\"%d\\n\", combine(x, x + 1)); return 0;", expect: ["12"] }
}
