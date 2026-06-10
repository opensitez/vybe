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
    tgamma_factorial_5 => {
        body: "printf(\"%.0f\\n\", tgamma(6.0));\nreturn 0;",
        expect: ["120"]
    },
    lgamma_positive => {
        body: "printf(\"%d\\n\", lgamma(5.0) > 0.0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    erf_zero => {
        body: "printf(\"%.1f\\n\", erf(0.0));\nreturn 0;",
        expect: ["0.0"]
    },
    erf_large_approaches_one => {
        body: "printf(\"%d\\n\", erf(3.0) > 0.99 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    erfc_complement => {
        body: "double x = 1.0;\nprintf(\"%d\\n\", fabs(erf(x) + erfc(x) - 1.0) < 1e-10 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    j0_bessel => {
        body: "printf(\"%d\\n\", fabs(j0(0.0) - 1.0) < 1e-10 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    nextafter_increases => {
        body: "double x = 1.0;\ndouble next = nextafter(x, 2.0);\nprintf(\"%d\\n\", next > x ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    copysign_applies_sign => {
        body: "printf(\"%.1f\\n\", copysign(3.0, -1.0));\nreturn 0;",
        expect: ["-3.0"]
    },
    nan_macro_is_nan => {
        body: "double n = NAN;\nprintf(\"%d\\n\", isnan(n) ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    huge_val_is_inf => {
        body: "printf(\"%d\\n\", isinf(HUGE_VAL) ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    }
}
