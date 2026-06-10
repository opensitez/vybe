use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<limits.h>", "<float.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    limits_int_max => {
        body: "printf(\"%d\\n\", INT_MAX);\nreturn 0;",
        expect: ["2147483647"]
    },
    limits_int_min => {
        body: "printf(\"%d\\n\", INT_MIN);\nreturn 0;",
        expect: ["-2147483648"]
    },
    limits_uint_max => {
        body: "printf(\"%u\\n\", UINT_MAX);\nreturn 0;",
        expect: ["4294967295"]
    },
    limits_char_max => {
        body: "printf(\"%d\\n\", CHAR_MAX);\nreturn 0;",
        expect: ["127"]
    },
    limits_char_min => {
        body: "printf(\"%d\\n\", CHAR_MIN);\nreturn 0;",
        expect: ["-128"]
    },
    limits_shrt_max => {
        body: "printf(\"%d\\n\", SHRT_MAX);\nreturn 0;",
        expect: ["32767"]
    },
    limits_long_max_positive => {
        body: "printf(\"%d\\n\", LONG_MAX > 0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    float_epsilon_positive => {
        body: "printf(\"%d\\n\", FLT_EPSILON > 0.0f ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    double_epsilon_positive => {
        body: "printf(\"%d\\n\", DBL_EPSILON > 0.0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    float_max_positive => {
        body: "printf(\"%d\\n\", FLT_MAX > 1.0f ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    flt_dig_reasonable => {
        body: "printf(\"%d\\n\", FLT_DIG >= 6 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    dbl_dig_reasonable => {
        body: "printf(\"%d\\n\", DBL_DIG >= 15 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    }
}
