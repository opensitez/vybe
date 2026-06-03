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
    enum_auto_increments_after_explicit_value => { declarations: "enum Level { LOW = 3, MEDIUM, HIGH };", body: "printf(\"%d %d %d\\n\", LOW, MEDIUM, HIGH);\nreturn 0;", expect: ["3 4 5"] },
    enum_value_can_be_compared_in_if => { declarations: "enum State { OFF, ON };", body: "enum State state = ON; if (state == ON) puts(\"on\"); else puts(\"off\");\nreturn 0;", expect: ["on"] },
    enum_value_can_drive_switch_case => { declarations: "enum State { OFF, ON };", body: "enum State state = OFF; switch (state) { case OFF: puts(\"off\"); break; case ON: puts(\"on\"); break; }\nreturn 0;", expect: ["off"] },
    typedef_enum_value_can_be_declared_without_enum_keyword => { declarations: "typedef enum { RED, GREEN, BLUE } Color;", body: "Color color = BLUE; printf(\"%d\\n\", color);\nreturn 0;", expect: ["2"] },
    enum_constant_can_be_used_in_array_index => { declarations: "enum Slot { FIRST, SECOND, THIRD };", body: "int values[3] = {10, 20, 30}; printf(\"%d\\n\", values[SECOND]);\nreturn 0;", expect: ["20"] },
    enum_constants_share_integer_semantics => { declarations: "enum Sign { NEG = -1, ZERO = 0, POS = 1 };", body: "printf(\"%d\\n\", NEG + POS);\nreturn 0;", expect: ["0"] },
    enum_value_can_be_assigned_from_constant => { declarations: "enum Mode { A = 4, B = 9 };", body: "enum Mode mode = B; printf(\"%d\\n\", mode);\nreturn 0;", expect: ["9"] },
    enum_constants_can_initialize_struct_field => { declarations: "enum State { IDLE, RUNNING }; struct Task { enum State state; };", body: "struct Task task = {RUNNING}; printf(\"%d\\n\", task.state);\nreturn 0;", expect: ["1"] },
    enum_constants_can_start_from_negative_value => { declarations: "enum Delta { DOWN = -2, SAME, UP };", body: "printf(\"%d %d %d\\n\", DOWN, SAME, UP);\nreturn 0;", expect: ["-2 -1 0"] },
    enum_variable_can_be_reassigned => { declarations: "enum State { OFF, ON };", body: "enum State state = OFF; state = ON; printf(\"%d\\n\", state);\nreturn 0;", expect: ["1"] },
    enum_can_be_passed_to_function => { declarations: "enum State { OFF, ON }; int is_on(enum State state) { return state == ON; }", body: "printf(\"%d\\n\", is_on(ON));\nreturn 0;", expect: ["1"] },
    enum_can_be_returned_from_function => { declarations: "enum Level { LOW, HIGH }; enum Level pick(void) { return HIGH; }", body: "printf(\"%d\\n\", pick());\nreturn 0;", expect: ["1"] },
    enum_array_can_store_multiple_constants => { declarations: "enum Digit { ZERO, ONE, TWO };", body: "enum Digit digits[3] = {ZERO, ONE, TWO}; printf(\"%d %d %d\\n\", digits[0], digits[1], digits[2]);\nreturn 0;", expect: ["0 1 2"] },
    enum_with_sparse_values_keeps_explicit_gaps => { declarations: "enum Code { OK = 200, MISSING = 404, FAIL = 500 };", body: "printf(\"%d %d %d\\n\", OK, MISSING, FAIL);\nreturn 0;", expect: ["200 404 500"] },
    enum_value_can_participate_in_ternary => { declarations: "enum Light { RED, GREEN };", body: "enum Light light = GREEN; puts(light == GREEN ? \"go\" : \"stop\");\nreturn 0;", expect: ["go"] },
    enum_variable_can_be_incremented_as_integer => { declarations: "enum Count { ZERO, ONE, TWO };", body: "enum Count count = ZERO; count = count + 2; printf(\"%d\\n\", count);\nreturn 0;", expect: ["2"] },
    enum_constant_can_be_used_as_case_label_after_explicit_base => { declarations: "enum Token { START = 10, END = 20 };", body: "int token = END; switch (token) { case START: puts(\"start\"); break; case END: puts(\"end\"); break; }\nreturn 0;", expect: ["end"] },
    enum_typedef_and_variable_can_share_namespace_rules => { declarations: "typedef enum { APPLE, PEAR } Fruit;", body: "Fruit fruit = PEAR; printf(\"%d\\n\", fruit);\nreturn 0;", expect: ["1"] },
    enum_inside_struct_can_be_read_via_field => { declarations: "enum State { OFF, ON }; struct Device { enum State state; };", body: "struct Device device = {ON}; printf(\"%d\\n\", device.state);\nreturn 0;", expect: ["1"] },
    enum_expression_can_feed_array_size_like_constant_usage => { declarations: "enum Size { LEN = 3 };", body: "int values[LEN] = {1, 2, 3}; printf(\"%d\\n\", values[2]);\nreturn 0;", expect: ["3"] },
    enum_comparison_between_different_named_constants_works_as_ints => { declarations: "enum A { A0 = 0 }; enum B { B0 = 0 };", body: "printf(\"%d\\n\", A0 == B0);\nreturn 0;", expect: ["1"] },
    enum_constant_can_initialize_global_variable => { declarations: "enum State { OFF, ON }; enum State current = ON;", body: "printf(\"%d\\n\", current);\nreturn 0;", expect: ["1"] },
    enum_field_can_be_updated_through_pointer_to_struct => { declarations: "enum State { OFF, ON }; struct Device { enum State state; };", body: "struct Device device = {OFF}; struct Device *p = &device; p->state = ON; printf(\"%d\\n\", device.state);\nreturn 0;", expect: ["1"] },
    enum_value_can_be_printed_with_decimal_format => { declarations: "enum Number { TEN = 10 };", body: "printf(\"%d\\n\", TEN);\nreturn 0;", expect: ["10"] },
    enum_constants_can_be_used_in_arithmetic_expression => { declarations: "enum Number { TWO = 2, THREE = 3 };", body: "printf(\"%d\\n\", TWO * THREE);\nreturn 0;", expect: ["6"] }
}