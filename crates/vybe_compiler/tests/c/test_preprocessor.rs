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
    object_like_define_substitutes_integer_constant => { declarations: "#define ANSWER 42", body: "printf(\"%d\\n\", ANSWER);\nreturn 0;", expect: ["42"] },
    object_like_define_can_be_used_in_expression => { declarations: "#define OFFSET 5", body: "printf(\"%d\\n\", OFFSET + 3);\nreturn 0;", expect: ["8"] },
    function_like_macro_can_add_values => { declarations: "#define ADD(a, b) ((a) + (b))", body: "printf(\"%d\\n\", ADD(3, 4));\nreturn 0;", expect: ["7"] },
    function_like_macro_can_preserve_precedence_with_parentheses => { declarations: "#define DOUBLE(x) ((x) * 2)", body: "printf(\"%d\\n\", DOUBLE(1 + 2));\nreturn 0;", expect: ["6"] },
    nested_macros_can_expand_in_order => { declarations: "#define BASE 10\n#define TWICE(x) ((x) * 2)", body: "printf(\"%d\\n\", TWICE(BASE));\nreturn 0;", expect: ["20"] },
    macro_can_produce_string_literal => { declarations: "#define NAME \"vybe\"", body: "puts(NAME);\nreturn 0;", expect: ["vybe"] },
    macro_can_wrap_comparison_expression => { declarations: "#define IS_POSITIVE(x) ((x) > 0)", body: "printf(\"%d\\n\", IS_POSITIVE(4));\nreturn 0;", expect: ["1"] },
    macro_can_be_reused_multiple_times => { declarations: "#define STEP 3", body: "printf(\"%d %d\\n\", STEP, STEP + STEP);\nreturn 0;", expect: ["3 6"] },
    macro_can_operate_on_character_literals => { declarations: "#define NEXT(c) ((c) + 1)", body: "printf(\"%c\\n\", NEXT('a'));\nreturn 0;", expect: ["b"] },
    macro_can_expand_inside_array_initializer => { declarations: "#define VALUE 7", body: "int arr[2] = {VALUE, VALUE + 1};\nprintf(\"%d %d\\n\", arr[0], arr[1]);\nreturn 0;", expect: ["7 8"] },
    macro_can_be_undefined_and_redefined => { declarations: "#define VALUE 3\n#undef VALUE\n#define VALUE 9", body: "printf(\"%d\\n\", VALUE);\nreturn 0;", expect: ["9"] },
    macro_can_expand_to_parenthesized_product => { declarations: "#define AREA(w, h) ((w) * (h))", body: "printf(\"%d\\n\", AREA(3, 4));\nreturn 0;", expect: ["12"] },
    macro_can_expand_to_boolean_condition => { declarations: "#define SHOULD_RUN 1", body: "if (SHOULD_RUN) puts(\"run\"); else puts(\"stop\");\nreturn 0;", expect: ["run"] },
    macro_can_expand_to_char_pointer => { declarations: "#define TEXT \"hello\"", body: "puts(TEXT);\nreturn 0;", expect: ["hello"] },
    macro_can_be_used_in_for_loop_bound => { declarations: "#define COUNT 3", body: "for (int i = 0; i < COUNT; i++) printf(\"%d\\n\", i);\nreturn 0;", expect: ["0", "1", "2"] },
    macro_can_expand_to_shift_expression => { declarations: "#define MASK (1 << 3)", body: "printf(\"%d\\n\", MASK);\nreturn 0;", expect: ["8"] },
    macro_argument_parentheses_prevent_precedence_bug => { declarations: "#define SQUARE(x) ((x) * (x))", body: "printf(\"%d\\n\", SQUARE(1 + 2));\nreturn 0;", expect: ["9"] },
    macro_can_expand_to_double_literal => { declarations: "#define PI_HALF 1.57079632679", body: "printf(\"%.2f\\n\", PI_HALF);\nreturn 0;", expect: ["1.57"] },
    include_stdio_provides_printf_and_puts => { declarations: "", body: "puts(\"stdio\");\nprintf(\"%d\\n\", 5);\nreturn 0;", expect: ["stdio", "5"] },
    macro_can_expand_inside_switch_case_value => { declarations: "#define MATCH 2", body: "int x = 2;\nswitch (x) { case MATCH: puts(\"hit\"); break; default: puts(\"miss\"); }\nreturn 0;", expect: ["hit"] },
    macro_can_expand_to_array_length_constant => { declarations: "#define LEN 4", body: "int arr[LEN] = {1, 2, 3, 4};\nprintf(\"%d\\n\", arr[LEN - 1]);\nreturn 0;", expect: ["4"] },
    macro_can_expand_to_cast_expression => { declarations: "#define HALF(x) ((double)(x) / 2.0)", body: "printf(\"%.1f\\n\", HALF(7));\nreturn 0;", expect: ["3.5"] },
    macro_can_expand_to_stringized_constant_like_value => { declarations: "#define MODE \"debug\"", body: "puts(MODE);\nreturn 0;", expect: ["debug"] },
    macro_can_chain_other_macros => { declarations: "#define ONE 1\n#define TWO (ONE + ONE)", body: "printf(\"%d\\n\", TWO);\nreturn 0;", expect: ["2"] },
    macro_can_define_character_constant => { declarations: "#define LETTER 'Z'", body: "printf(\"%c\\n\", LETTER);\nreturn 0;", expect: ["Z"] }
}