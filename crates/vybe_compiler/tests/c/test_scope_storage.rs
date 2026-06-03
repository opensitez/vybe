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
    global_variable_can_be_read_in_main => { declarations: "int global_value = 7;", body: "printf(\"%d\\n\", global_value);\nreturn 0;", expect: ["7"] },
    global_variable_can_be_written_in_main => { declarations: "int global_value = 7;", body: "global_value = 9;\nprintf(\"%d\\n\", global_value);\nreturn 0;", expect: ["9"] },
    local_variable_can_shadow_global_name => { declarations: "int value = 5;", body: "int value = 9;\nprintf(\"%d\\n\", value);\nreturn 0;", expect: ["9"] },
    shadowed_global_remains_unchanged_after_inner_block => { declarations: "int value = 5;", body: "{ int value = 9; printf(\"%d\\n\", value); }\nprintf(\"%d\\n\", value);\nreturn 0;", expect: ["9", "5"] },
    block_scope_variable_is_distinct_from_outer_scope => { declarations: "", body: "int x = 1; { int x = 2; printf(\"%d\\n\", x); } printf(\"%d\\n\", x);\nreturn 0;", expect: ["2", "1"] },
    file_scope_function_can_read_file_scope_variable => { declarations: "int base = 10; int add_base(int x) { return base + x; }", body: "printf(\"%d\\n\", add_base(4));\nreturn 0;", expect: ["14"] },
    static_global_variable_can_be_read => { declarations: "static int hidden = 12;", body: "printf(\"%d\\n\", hidden);\nreturn 0;", expect: ["12"] },
    static_local_variable_persists_across_calls => { declarations: "int tick(void) { static int n = 0; n++; return n; }", body: "printf(\"%d\\n\", tick());\nprintf(\"%d\\n\", tick());\nreturn 0;", expect: ["1", "2"] },
    extern_declaration_can_reference_same_translation_unit_global => { declarations: "int shared = 21; extern int shared;", body: "printf(\"%d\\n\", shared);\nreturn 0;", expect: ["21"] },
    inner_block_can_reuse_name_without_affecting_outer_assignment => { declarations: "", body: "int x = 3; { int x = 4; x += 1; printf(\"%d\\n\", x); } printf(\"%d\\n\", x);\nreturn 0;", expect: ["5", "3"] },
    for_loop_variable_scope_does_not_escape_body_use => { declarations: "", body: "for (int i = 0; i < 2; i++) printf(\"%d\\n\", i);\nreturn 0;", expect: ["0", "1"] },
    function_parameter_shadows_global_name => { declarations: "int value = 3; int show(int value) { return value; }", body: "printf(\"%d %d\\n\", show(8), value);\nreturn 0;", expect: ["8 3"] },
    nested_block_can_access_outer_variable_when_unshadowed => { declarations: "", body: "int x = 4; { printf(\"%d\\n\", x); }\nreturn 0;", expect: ["4"] },
    nested_block_shadowing_can_use_both_levels => { declarations: "", body: "int x = 4; { int y = x + 1; int x = 9; printf(\"%d %d\\n\", x, y); }\nreturn 0;", expect: ["9 5"] },
    static_function_can_be_called_from_main => { declarations: "static int twice(int x) { return x * 2; }", body: "printf(\"%d\\n\", twice(5));\nreturn 0;", expect: ["10"] },
    local_variable_lifetime_is_per_call => { declarations: "int next(int x) { int y = x + 1; return y; }", body: "printf(\"%d %d\\n\", next(1), next(1));\nreturn 0;", expect: ["2 2"] },
    static_local_and_regular_local_can_coexist => { declarations: "int sample(void) { static int s = 0; int t = 5; s++; return s + t; }", body: "printf(\"%d\\n\", sample());\nprintf(\"%d\\n\", sample());\nreturn 0;", expect: ["6", "7"] },
    global_array_can_be_read_from_main => { declarations: "int values[3] = {2, 4, 6};", body: "printf(\"%d\\n\", values[1]);\nreturn 0;", expect: ["4"] },
    global_struct_can_be_read_from_main => { declarations: "struct Pair { int a; int b; }; struct Pair pair = {3, 5};", body: "printf(\"%d\\n\", pair.b);\nreturn 0;", expect: ["5"] },
    file_scope_pointer_can_reference_global => { declarations: "int value = 8; int *ptr = &value;", body: "printf(\"%d\\n\", *ptr);\nreturn 0;", expect: ["8"] },
    shadowed_name_in_if_block_does_not_change_outer_value => { declarations: "", body: "int x = 1; if (1) { int x = 7; printf(\"%d\\n\", x); } printf(\"%d\\n\", x);\nreturn 0;", expect: ["7", "1"] },
    nested_scope_can_mutate_outer_variable_when_not_shadowed => { declarations: "", body: "int x = 1; { x = 5; } printf(\"%d\\n\", x);\nreturn 0;", expect: ["5"] },
    static_array_local_persists_written_element_between_calls => { declarations: "int second(void) { static int values[2] = {1, 2}; values[1] += 1; return values[1]; }", body: "printf(\"%d\\n\", second());\nprintf(\"%d\\n\", second());\nreturn 0;", expect: ["3", "4"] },
    extern_function_prototype_can_precede_definition => { declarations: "extern int add(int, int); int add(int a, int b) { return a + b; }", body: "printf(\"%d\\n\", add(2, 3));\nreturn 0;", expect: ["5"] },
    inner_scope_can_declare_same_typedef_name_shadow_independently => { declarations: "typedef int Number;", body: "Number a = 3; { typedef int Number; Number b = 4; printf(\"%d\\n\", b); } printf(\"%d\\n\", a);\nreturn 0;", expect: ["4", "3"] }
}