use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<math.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    sin_of_zero => { body: "printf(\"%.3f\\n\", sin(0.0));\nreturn 0;", expect: ["0.000"] },
    sin_of_pi_over_two => { body: "printf(\"%.3f\\n\", sin(M_PI / 2.0));\nreturn 0;", expect: ["1.000"] },
    sin_of_negative_pi_over_two => { body: "printf(\"%.3f\\n\", sin(-M_PI / 2.0));\nreturn 0;", expect: ["-1.000"] },
    sin_of_pi => { body: "printf(\"%.3f\\n\", sin(M_PI));\nreturn 0;", expect: ["0.000"] },
    cos_of_zero => { body: "printf(\"%.3f\\n\", cos(0.0));\nreturn 0;", expect: ["1.000"] },
    cos_of_pi => { body: "printf(\"%.3f\\n\", cos(M_PI));\nreturn 0;", expect: ["-1.000"] },
    cos_of_pi_over_two => { body: "printf(\"%.3f\\n\", cos(M_PI / 2.0));\nreturn 0;", expect: ["0.000"] },
    tan_of_zero => { body: "printf(\"%.3f\\n\", tan(0.0));\nreturn 0;", expect: ["0.000"] },
    tan_of_pi_over_four => { body: "printf(\"%.3f\\n\", tan(M_PI / 4.0));\nreturn 0;", expect: ["1.000"] },
    asin_of_zero => { body: "printf(\"%.3f\\n\", asin(0.0));\nreturn 0;", expect: ["0.000"] },
    asin_of_one => { body: "printf(\"%.3f\\n\", asin(1.0));\nreturn 0;", expect: ["1.571"] },
    acos_of_one => { body: "printf(\"%.3f\\n\", acos(1.0));\nreturn 0;", expect: ["0.000"] },
    acos_of_zero => { body: "printf(\"%.3f\\n\", acos(0.0));\nreturn 0;", expect: ["1.571"] },
    atan_of_zero => { body: "printf(\"%.3f\\n\", atan(0.0));\nreturn 0;", expect: ["0.000"] },
    atan_of_one => { body: "printf(\"%.3f\\n\", atan(1.0));\nreturn 0;", expect: ["0.785"] },
    atan_of_negative_one => { body: "printf(\"%.3f\\n\", atan(-1.0));\nreturn 0;", expect: ["-0.785"] },
    atan2_first_quadrant => { body: "printf(\"%.3f\\n\", atan2(1.0, 1.0));\nreturn 0;", expect: ["0.785"] },
    atan2_second_quadrant => { body: "printf(\"%.3f\\n\", atan2(1.0, -1.0));\nreturn 0;", expect: ["2.356"] },
    atan2_third_quadrant => { body: "printf(\"%.3f\\n\", atan2(-1.0, -1.0));\nreturn 0;", expect: ["-2.356"] },
    atan2_fourth_quadrant => { body: "printf(\"%.3f\\n\", atan2(-1.0, 1.0));\nreturn 0;", expect: ["-0.785"] },
    atan2_positive_y_axis => { body: "printf(\"%.3f\\n\", atan2(1.0, 0.0));\nreturn 0;", expect: ["1.571"] },
    exp_of_zero => { body: "printf(\"%.3f\\n\", exp(0.0));\nreturn 0;", expect: ["1.000"] },
    exp_of_one => { body: "printf(\"%.3f\\n\", exp(1.0));\nreturn 0;", expect: ["2.718"] },
    log_of_one => { body: "printf(\"%.3f\\n\", log(1.0));\nreturn 0;", expect: ["0.000"] },
    log_of_e => { body: "printf(\"%.3f\\n\", log(M_E));\nreturn 0;", expect: ["1.000"] },
    log10_of_one => { body: "printf(\"%.3f\\n\", log10(1.0));\nreturn 0;", expect: ["0.000"] },
    log10_of_thousand => { body: "printf(\"%.3f\\n\", log10(1000.0));\nreturn 0;", expect: ["3.000"] },
    exp_and_log_cancel => { body: "printf(\"%.3f\\n\", log(exp(2.0)));\nreturn 0;", expect: ["2.000"] },
    trig_identity_sine_squared_plus_cosine_squared => { body: "double x = 0.75;\nprintf(\"%.3f\\n\", sin(x) * sin(x) + cos(x) * cos(x));\nreturn 0;", expect: ["1.000"] },
    atan_tan_roundtrip_small_angle => { body: "double x = 0.25;\nprintf(\"%.3f\\n\", atan(tan(x)));\nreturn 0;", expect: ["0.250"] }
}