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
    typedef_alias_for_int_can_declare_variable => { declarations: "typedef int Number;", body: "Number value = 7; printf(\"%d\\n\", value);\nreturn 0;", expect: ["7"] },
    typedef_alias_for_pointer_can_reference_value => { declarations: "typedef int *IntPtr;", body: "int value = 8; IntPtr ptr = &value; printf(\"%d\\n\", *ptr);\nreturn 0;", expect: ["8"] },
    typedef_alias_for_struct_can_be_used_without_struct_keyword => { declarations: "typedef struct { int x; int y; } Point;", body: "Point point = {3, 4}; printf(\"%d %d\\n\", point.x, point.y);\nreturn 0;", expect: ["3 4"] },
    typedef_alias_for_char_pointer_can_hold_string_literal => { declarations: "typedef char *Text;", body: "Text text = \"vybe\"; puts(text);\nreturn 0;", expect: ["vybe"] },
    typedef_alias_for_array_element_pointer_can_index => { declarations: "typedef int *IntPtr;", body: "int values[3] = {1, 2, 3}; IntPtr ptr = values; printf(\"%d\\n\", ptr[2]);\nreturn 0;", expect: ["3"] },
    typedef_alias_for_function_pointer_can_invoke_target => { declarations: "typedef int (*Unary)(int); int add_one(int x) { return x + 1; }", body: "Unary fn = add_one; printf(\"%d\\n\", fn(4));\nreturn 0;", expect: ["5"] },
    typedef_alias_can_be_used_in_function_signature => { declarations: "typedef int Number; Number double_it(Number x) { return x * 2; }", body: "printf(\"%d\\n\", double_it(6));\nreturn 0;", expect: ["12"] },
    typedef_alias_for_unsigned_can_print_with_unsigned_format => { declarations: "typedef unsigned int Flags;", body: "Flags flags = 7u; printf(\"%u\\n\", flags);\nreturn 0;", expect: ["7"] },
    typedef_alias_for_nested_struct_can_copy_value => { declarations: "typedef struct { int a; int b; } Pair;", body: "Pair first = {1, 2}; Pair second = first; printf(\"%d %d\\n\", second.a, second.b);\nreturn 0;", expect: ["1 2"] },
    typedef_alias_for_function_pointer_parameter_can_be_passed => { declarations: "typedef int (*Unary)(int); int apply(Unary fn, int value) { return fn(value); } int square(int x) { return x * x; }", body: "printf(\"%d\\n\", apply(square, 5));\nreturn 0;", expect: ["25"] },
    typedef_alias_can_shadow_builtin_type_name_locally => { declarations: "typedef int Count;", body: "Count count = 3; printf(\"%d\\n\", count);\nreturn 0;", expect: ["3"] },
    typedef_alias_for_pointer_to_struct_can_follow_field => { declarations: "typedef struct { int x; } Point; typedef Point *PointPtr;", body: "Point point = {9}; PointPtr ptr = &point; printf(\"%d\\n\", ptr->x);\nreturn 0;", expect: ["9"] },
    typedef_alias_can_be_used_for_array_length_enum_combo => { declarations: "typedef int Number; enum { LEN = 3 };", body: "Number values[LEN] = {2, 4, 6}; printf(\"%d\\n\", values[2]);\nreturn 0;", expect: ["6"] },
    typedef_alias_for_double_preserves_fraction => { declarations: "typedef double Real;", body: "Real value = 2.5; printf(\"%.1f\\n\", value);\nreturn 0;", expect: ["2.5"] },
    typedef_alias_for_char_can_hold_letter => { declarations: "typedef char Letter;", body: "Letter letter = 'Q'; printf(\"%c\\n\", letter);\nreturn 0;", expect: ["Q"] },
    typedef_alias_for_const_pointer_like_shape_can_read_string => { declarations: "typedef char *Text;", body: "Text text = \"hello\"; printf(\"%c\\n\", text[1]);\nreturn 0;", expect: ["e"] },
    typedef_struct_and_function_pointer_can_coexist => { declarations: "typedef struct { int value; } Box; typedef int (*Getter)(Box *); int get(Box *box) { return box->value; }", body: "Box box = {7}; Getter getter = get; printf(\"%d\\n\", getter(&box));\nreturn 0;", expect: ["7"] },
    typedef_alias_for_array_pointer_can_reference_first_element => { declarations: "typedef int *IntPtr;", body: "int arr[2] = {5, 6}; IntPtr ptr = &arr[0]; printf(\"%d\\n\", *ptr);\nreturn 0;", expect: ["5"] },
    typedef_alias_can_be_used_in_nested_scope => { declarations: "typedef int Number;", body: "{ Number value = 11; printf(\"%d\\n\", value); }\nreturn 0;", expect: ["11"] },
    typedef_alias_for_function_return_type_can_be_used => { declarations: "typedef int Number; Number answer(void) { return 42; }", body: "printf(\"%d\\n\", answer());\nreturn 0;", expect: ["42"] },
    typedef_alias_for_enum_can_declare_variable => { declarations: "typedef enum { NO, YES } Boolish;", body: "Boolish value = YES; printf(\"%d\\n\", value);\nreturn 0;", expect: ["1"] },
    typedef_alias_for_union_can_store_value => { declarations: "typedef union { int i; char c; } Data;", body: "Data data; data.i = 65; printf(\"%d\\n\", data.i);\nreturn 0;", expect: ["65"] },
    typedef_alias_for_pointer_to_function_returning_pointer_can_call => { declarations: "typedef int *(*IdFn)(int *); int *identity(int *p) { return p; }", body: "int value = 12; IdFn fn = identity; printf(\"%d\\n\", *fn(&value));\nreturn 0;", expect: ["12"] },
    typedef_alias_for_signed_char_can_print_ascii => { declarations: "typedef signed char Small;", body: "Small value = 65; printf(\"%d\\n\", value);\nreturn 0;", expect: ["65"] },
    typedef_alias_for_array_of_function_pointers_can_dispatch => { declarations: "typedef int (*Unary)(int); int add_one(int x) { return x + 1; } int double_it(int x) { return x * 2; }", body: "Unary ops[2] = {add_one, double_it}; printf(\"%d %d\\n\", ops[0](3), ops[1](3));\nreturn 0;", expect: ["4 6"] }
}