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
    logical_and_true_true => { declarations: "", body: "printf(\"%d\\n\", 1 && 2);\nreturn 0;", expect: ["1"] },
    logical_and_true_false => { declarations: "", body: "printf(\"%d\\n\", 1 && 0);\nreturn 0;", expect: ["0"] },
    logical_or_false_true => { declarations: "", body: "printf(\"%d\\n\", 0 || 9);\nreturn 0;", expect: ["1"] },
    logical_or_false_false => { declarations: "", body: "printf(\"%d\\n\", 0 || 0);\nreturn 0;", expect: ["0"] },
    logical_not_zero_is_true => { declarations: "", body: "printf(\"%d\\n\", !0);\nreturn 0;", expect: ["1"] },
    logical_not_nonzero_is_false => { declarations: "", body: "printf(\"%d\\n\", !5);\nreturn 0;", expect: ["0"] },
    less_than_comparison_yields_true => { declarations: "", body: "printf(\"%d\\n\", 2 < 3);\nreturn 0;", expect: ["1"] },
    greater_than_comparison_yields_false => { declarations: "", body: "printf(\"%d\\n\", 2 > 3);\nreturn 0;", expect: ["0"] },
    equality_of_same_values_is_true => { declarations: "", body: "printf(\"%d\\n\", 7 == 7);\nreturn 0;", expect: ["1"] },
    inequality_of_same_values_is_false => { declarations: "", body: "printf(\"%d\\n\", 7 != 7);\nreturn 0;", expect: ["0"] },
    greater_equal_of_equal_values_is_true => { declarations: "", body: "printf(\"%d\\n\", 7 >= 7);\nreturn 0;", expect: ["1"] },
    less_equal_of_greater_value_is_false => { declarations: "", body: "printf(\"%d\\n\", 9 <= 7);\nreturn 0;", expect: ["0"] },
    negative_value_is_truthy_in_if => { declarations: "", body: "if (-1) puts(\"true\"); else puts(\"false\");\nreturn 0;", expect: ["true"] },
    zero_is_falsey_in_if => { declarations: "", body: "if (0) puts(\"true\"); else puts(\"false\");\nreturn 0;", expect: ["false"] },
    logical_and_short_circuits_false_left_side => { declarations: "int x = 0;", body: "if (0 && ++x) puts(\"bad\");\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["0"] },
    logical_or_short_circuits_true_left_side => { declarations: "int x = 0;", body: "if (1 || ++x) puts(\"ok\");\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["ok", "0"] },
    logical_and_allows_right_side_when_left_true => { declarations: "int x = 0;", body: "if (1 && ++x) puts(\"ok\");\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["ok", "1"] },
    logical_or_evaluates_right_side_when_left_false => { declarations: "int x = 0;", body: "if (0 || ++x) puts(\"ok\");\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["ok", "1"] },
    logical_and_precedence_over_or => { declarations: "", body: "printf(\"%d\\n\", 0 || 1 && 0);\nreturn 0;", expect: ["0"] },
    parentheses_override_logical_precedence => { declarations: "", body: "printf(\"%d\\n\", (0 || 1) && 0);\nreturn 0;", expect: ["0"] },
    equality_compares_after_addition => { declarations: "", body: "printf(\"%d\\n\", 1 + 2 == 3);\nreturn 0;", expect: ["1"] },
    chained_comparison_uses_integer_result_of_first_comparison => { declarations: "", body: "printf(\"%d\\n\", (1 < 2) < 3);\nreturn 0;", expect: ["1"] },
    logical_not_applies_before_equality => { declarations: "", body: "printf(\"%d\\n\", !1 == 0);\nreturn 0;", expect: ["1"] },
    logical_or_can_avoid_division_by_zero => { declarations: "", body: "if (1 || (10 / 0)) puts(\"safe\"); else puts(\"bad\");\nreturn 0;", expect: ["safe"] },
    logical_and_can_avoid_division_by_zero => { declarations: "", body: "if (0 && (10 / 0)) puts(\"bad\"); else puts(\"safe\");\nreturn 0;", expect: ["safe"] },
    double_comparison_is_true => { declarations: "", body: "printf(\"%d\\n\", 3.5 > 3.4);\nreturn 0;", expect: ["1"] },
    comparison_result_can_be_added => { declarations: "", body: "printf(\"%d\\n\", (3 > 2) + (2 > 3));\nreturn 0;", expect: ["1"] },
    nonzero_logical_and_returns_one_not_operand => { declarations: "", body: "printf(\"%d\\n\", 2 && 4);\nreturn 0;", expect: ["1"] },
    nonzero_logical_or_returns_one_not_operand => { declarations: "", body: "printf(\"%d\\n\", 2 || 4);\nreturn 0;", expect: ["1"] },
    equality_after_logical_not_zero => { declarations: "", body: "printf(\"%d\\n\", (!0) == 1);\nreturn 0;", expect: ["1"] }
}
