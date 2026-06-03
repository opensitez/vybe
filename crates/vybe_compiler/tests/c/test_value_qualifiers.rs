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
    const_int_variable_can_be_read => { declarations: "const int value = 7;", body: "printf(\"%d\\n\", value);\nreturn 0;", expect: ["7"] },
    const_double_variable_can_be_read => { declarations: "const double value = 2.5;", body: "printf(\"%.1f\\n\", value);\nreturn 0;", expect: ["2.5"] },
    pointer_to_const_can_read_target => { declarations: "int value = 9; const int *p = &value;", body: "printf(\"%d\\n\", *p);\nreturn 0;", expect: ["9"] },
    const_pointer_can_keep_same_target_for_reads => { declarations: "int value = 9; int *const p = &value;", body: "printf(\"%d\\n\", *p);\nreturn 0;", expect: ["9"] },
    const_pointer_can_write_through_nonconst_pointee => { declarations: "int value = 9; int *const p = &value;", body: "*p = 12; printf(\"%d\\n\", value);\nreturn 0;", expect: ["12"] },
    const_array_can_be_indexed_for_read => { declarations: "const int values[3] = {1, 2, 3};", body: "printf(\"%d\\n\", values[2]);\nreturn 0;", expect: ["3"] },
    const_char_pointer_can_read_string_literal => { declarations: "const char *text = \"vybe\";", body: "puts(text);\nreturn 0;", expect: ["vybe"] },
    static_global_value_is_available_in_main => { declarations: "static int value = 11;", body: "printf(\"%d\\n\", value);\nreturn 0;", expect: ["11"] },
    static_local_counter_persists_across_calls => { declarations: "int next_id(void) { static int id = 100; return ++id; }", body: "printf(\"%d\\n\", next_id());\nprintf(\"%d\\n\", next_id());\nreturn 0;", expect: ["101", "102"] },
    volatile_int_variable_can_be_read => { declarations: "volatile int value = 13;", body: "printf(\"%d\\n\", value);\nreturn 0;", expect: ["13"] },
    volatile_pointer_target_can_be_read => { declarations: "int value = 14; volatile int *p = &value;", body: "printf(\"%d\\n\", *p);\nreturn 0;", expect: ["14"] },
    const_struct_field_container_can_be_read => { declarations: "struct Pair { int a; int b; }; const struct Pair pair = {3, 4};", body: "printf(\"%d %d\\n\", pair.a, pair.b);\nreturn 0;", expect: ["3 4"] },
    static_array_local_persists_updates_between_calls => { declarations: "int next_slot(void) { static int values[2] = {1, 2}; values[0] += 1; return values[0]; }", body: "printf(\"%d\\n\", next_slot());\nprintf(\"%d\\n\", next_slot());\nreturn 0;", expect: ["2", "3"] },
    const_enum_like_object_can_drive_condition => { declarations: "const int enabled = 1;", body: "if (enabled) puts(\"on\"); else puts(\"off\");\nreturn 0;", expect: ["on"] },
    static_function_can_return_constant_value => { declarations: "static int answer(void) { return 42; }", body: "printf(\"%d\\n\", answer());\nreturn 0;", expect: ["42"] },
    const_pointer_to_const_can_read_character => { declarations: "const char *const text = \"hi\";", body: "printf(\"%c\\n\", text[1]);\nreturn 0;", expect: ["i"] },
    volatile_struct_field_container_can_be_read => { declarations: "struct Pair { int a; int b; }; volatile struct Pair pair = {5, 6};", body: "printf(\"%d %d\\n\", pair.a, pair.b);\nreturn 0;", expect: ["5 6"] },
    const_pointer_parameter_can_be_passed_to_function => { declarations: "int first(const int *values) { return values[0]; }", body: "const int values[2] = {7, 8}; printf(\"%d\\n\", first(values));\nreturn 0;", expect: ["7"] },
    static_local_and_global_names_can_coexist => { declarations: "int value = 3; int sample(void) { static int value = 5; return value; }", body: "printf(\"%d %d\\n\", sample(), value);\nreturn 0;", expect: ["5 3"] },
    volatile_local_can_be_updated_and_read => { declarations: "", body: "volatile int value = 4; value += 3; printf(\"%d\\n\", value);\nreturn 0;", expect: ["7"] },
    const_double_can_participate_in_expression => { declarations: "const double pi = 3.14;", body: "printf(\"%.2f\\n\", pi * 2.0);\nreturn 0;", expect: ["6.28"] },
    static_struct_local_persists_field_updates => { declarations: "int step(void) { static struct Pair { int a; } pair = {1}; pair.a += 1; return pair.a; }", body: "printf(\"%d\\n\", step());\nprintf(\"%d\\n\", step());\nreturn 0;", expect: ["2", "3"] },
    const_char_array_can_be_measured => { declarations: "const char text[] = \"abc\";", body: "printf(\"%d\\n\", (int)sizeof(text));\nreturn 0;", expect: ["4"] },
    static_pointer_can_reference_global_storage => { declarations: "int value = 9; static int *p = &value;", body: "printf(\"%d\\n\", *p);\nreturn 0;", expect: ["9"] },
    const_pointer_to_array_can_index_elements => { declarations: "int values[3] = {2, 4, 6}; int *const p = values;", body: "printf(\"%d\\n\", p[2]);\nreturn 0;", expect: ["6"] }
}