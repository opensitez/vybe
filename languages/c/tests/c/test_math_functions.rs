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
    pow_integer_exponent => {
        body: "printf(\"%.0f\\n\", pow(2.0, 10.0));\nreturn 0;",
        expect: ["1024"]
    },
    sqrt_perfect_square => {
        body: "printf(\"%.0f\\n\", sqrt(144.0));\nreturn 0;",
        expect: ["12"]
    },
    ceil_positive => {
        body: "printf(\"%.0f\\n\", ceil(3.1));\nreturn 0;",
        expect: ["4"]
    },
    floor_positive => {
        body: "printf(\"%.0f\\n\", floor(3.9));\nreturn 0;",
        expect: ["3"]
    },
    fmod_basic => {
        body: "printf(\"%.1f\\n\", fmod(10.5, 3.0));\nreturn 0;",
        expect: ["1.5"]
    },
    exp_euler_approx => {
        body: "printf(\"%.4f\\n\", exp(1.0));\nreturn 0;",
        expect: ["2.7183"]
    },
    log_natural => {
        body: "printf(\"%.4f\\n\", log(2.71828));\nreturn 0;",
        expect: ["1.0000"]
    },
    log10_base10 => {
        body: "printf(\"%.1f\\n\", log10(1000.0));\nreturn 0;",
        expect: ["3.0"]
    },
    sin_zero => {
        body: "printf(\"%.1f\\n\", sin(0.0));\nreturn 0;",
        expect: ["0.0"]
    },
    cos_zero => {
        body: "printf(\"%.1f\\n\", cos(0.0));\nreturn 0;",
        expect: ["1.0"]
    },
    atan2_quadrant => {
        body: "printf(\"%.4f\\n\", atan2(1.0, 1.0) * 4 / 3.14159);\nreturn 0;",
        expect: ["1.0000"]
    },
    pi_constant_approximation => {
        body: "printf(\"%.5f\\n\", 4.0 * atan(1.0));\nreturn 0;",
        expect: ["3.14159"]
    }
}
