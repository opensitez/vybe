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
    hypot_pythagorean => {
        body: "printf(\"%.1f\\n\", hypot(3.0, 4.0));\nreturn 0;",
        expect: ["5.0"]
    },
    cbrt_perfect_cube => {
        body: "printf(\"%.1f\\n\", cbrt(27.0));\nreturn 0;",
        expect: ["3.0"]
    },
    fmin_selects_smaller => {
        body: "printf(\"%.1f\\n\", fmin(3.0, 7.0));\nreturn 0;",
        expect: ["3.0"]
    },
    fmax_selects_larger => {
        body: "printf(\"%.1f\\n\", fmax(3.0, 7.0));\nreturn 0;",
        expect: ["7.0"]
    },
    fabs_removes_sign => {
        body: "printf(\"%.1f\\n\", fabs(-5.5));\nreturn 0;",
        expect: ["5.5"]
    },
    isnan_detects_nan => {
        body: "double nan_val = 0.0 / 0.0;\nprintf(\"%d\\n\", isnan(nan_val) ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    isinf_detects_infinity => {
        body: "double inf_val = 1.0 / 0.0;\nprintf(\"%d\\n\", isinf(inf_val) ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    isfinite_detects_finite => {
        body: "printf(\"%d\\n\", isfinite(3.14) ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    frexp_splits_mantissa_exponent => {
        body: "int e;\ndouble m = frexp(8.0, &e);\nprintf(\"%.4f %d\\n\", m, e);\nreturn 0;",
        expect: ["0.5000 4"]
    },
    ldexp_combines_mantissa_exponent => {
        body: "printf(\"%.1f\\n\", ldexp(0.5, 4));\nreturn 0;",
        expect: ["8.0"]
    },
    modf_splits_integer_fractional => {
        body: "double frac, intpart;\nfrac = modf(3.7, &intpart);\nprintf(\"%.1f %.1f\\n\", intpart, frac);\nreturn 0;",
        expect: ["3.0 0.7"]
    },
    log2_power_of_two => {
        body: "printf(\"%.1f\\n\", log2(8.0));\nreturn 0;",
        expect: ["3.0"]
    },
    exp2_basic => {
        body: "printf(\"%.1f\\n\", exp2(3.0));\nreturn 0;",
        expect: ["8.0"]
    },
    trunc_truncates_toward_zero => {
        body: "printf(\"%.1f %.1f\\n\", trunc(3.9), trunc(-3.9));\nreturn 0;",
        expect: ["3.0 -3.0"]
    },
    round_rounds_half_up => {
        body: "printf(\"%.1f %.1f\\n\", round(3.5), round(-3.5));\nreturn 0;",
        expect: ["4.0 -4.0"]
    }
}
