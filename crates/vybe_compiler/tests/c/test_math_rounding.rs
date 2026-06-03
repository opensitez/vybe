use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { includes: [$($include:expr),* $(,)?], body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&[$($include),*], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    sqrt_of_perfect_square => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.1f\\n\", sqrt(81.0));\nreturn 0;", expect: ["9.0"] },
    sqrt_of_fraction => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.3f\\n\", sqrt(0.25));\nreturn 0;", expect: ["0.500"] },
    fabs_of_negative_fraction => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.3f\\n\", fabs(-7.125));\nreturn 0;", expect: ["7.125"] },
    floor_of_positive_fraction => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.0f\\n\", floor(7.9));\nreturn 0;", expect: ["7"] },
    floor_of_negative_fraction => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.0f\\n\", floor(-7.1));\nreturn 0;", expect: ["-8"] },
    ceil_of_positive_fraction => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.0f\\n\", ceil(7.1));\nreturn 0;", expect: ["8"] },
    ceil_of_negative_fraction => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.0f\\n\", ceil(-7.9));\nreturn 0;", expect: ["-7"] },
    round_half_up_positive => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.0f\\n\", round(4.5));\nreturn 0;", expect: ["5"] },
    round_half_away_from_zero_negative => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.0f\\n\", round(-4.5));\nreturn 0;", expect: ["-5"] },
    round_small_fraction_down => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.0f\\n\", round(4.49));\nreturn 0;", expect: ["4"] },
    pow_integer_exponent => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.0f\\n\", pow(3.0, 4.0));\nreturn 0;", expect: ["81"] },
    pow_fractional_exponent => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.3f\\n\", pow(9.0, 0.5));\nreturn 0;", expect: ["3.000"] },
    pow_negative_base_even_exponent => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.0f\\n\", pow(-3.0, 2.0));\nreturn 0;", expect: ["9"] },
    pow_negative_base_odd_exponent => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.0f\\n\", pow(-3.0, 3.0));\nreturn 0;", expect: ["-27"] },
    fmod_basic_remainder => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.1f\\n\", fmod(17.0, 5.0));\nreturn 0;", expect: ["2.0"] },
    fmod_negative_dividend => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.1f\\n\", fmod(-17.0, 5.0));\nreturn 0;", expect: ["-2.0"] },
    fabs_preserves_zero => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.1f\\n\", fabs(0.0));\nreturn 0;", expect: ["0.0"] },
    math_constant_pi => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.3f\\n\", M_PI);\nreturn 0;", expect: ["3.142"] },
    math_constant_e => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.3f\\n\", M_E);\nreturn 0;", expect: ["2.718"] },
    math_constant_sqrt2 => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.3f\\n\", M_SQRT2);\nreturn 0;", expect: ["1.414"] },
    abs_positive_integer => { includes: ["<stdio.h>", "<stdlib.h>"], body: "printf(\"%d\\n\", abs(9));\nreturn 0;", expect: ["9"] },
    abs_negative_integer => { includes: ["<stdio.h>", "<stdlib.h>"], body: "printf(\"%d\\n\", abs(-9));\nreturn 0;", expect: ["9"] },
    labs_positive_long => { includes: ["<stdio.h>", "<stdlib.h>"], body: "printf(\"%d\\n\", labs(9001));\nreturn 0;", expect: ["9001"] },
    labs_negative_long => { includes: ["<stdio.h>", "<stdlib.h>"], body: "printf(\"%d\\n\", labs(-9001));\nreturn 0;", expect: ["9001"] },
    nested_rounding_expression => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.0f\\n\", ceil(fabs(-3.2) + 0.1));\nreturn 0;", expect: ["4"] },
    sqrt_then_square_restores_value => { includes: ["<stdio.h>", "<math.h>"], body: "double x = 49.0;\nprintf(\"%.0f\\n\", pow(sqrt(x), 2.0));\nreturn 0;", expect: ["49"] },
    floor_then_add_fraction => { includes: ["<stdio.h>", "<math.h>"], body: "double x = floor(8.9) + 0.25;\nprintf(\"%.2f\\n\", x);\nreturn 0;", expect: ["8.25"] },
    ceil_then_subtract_fraction => { includes: ["<stdio.h>", "<math.h>"], body: "double x = ceil(8.1) - 0.25;\nprintf(\"%.2f\\n\", x);\nreturn 0;", expect: ["8.75"] },
    round_expression_result => { includes: ["<stdio.h>", "<math.h>"], body: "double x = 2.2 + 2.3;\nprintf(\"%.0f\\n\", round(x));\nreturn 0;", expect: ["5"] },
    fmod_zero_remainder => { includes: ["<stdio.h>", "<math.h>"], body: "printf(\"%.1f\\n\", fmod(12.0, 3.0));\nreturn 0;", expect: ["0.0"] }
}