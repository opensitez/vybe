use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    int_to_float_promotion => {
        body: "int i = 5;\nfloat f = i;\nprintf(\"%.1f\\n\", f);\nreturn 0;",
        expect: ["5.0"]
    },
    float_to_int_truncates => {
        body: "float f = 3.9f;\nint i = f;\nprintf(\"%d\\n\", i);\nreturn 0;",
        expect: ["3"]
    },
    char_widened_to_int => {
        body: "char c = 100;\nint x = c + 100;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["200"]
    },
    int_narrowed_to_char => {
        body: "int n = 65;\nchar c = n;\nprintf(\"%c\\n\", c);\nreturn 0;",
        expect: ["A"]
    },
    double_to_int_truncates => {
        body: "double d = -2.7;\nint i = d;\nprintf(\"%d\\n\", i);\nreturn 0;",
        expect: ["-2"]
    },
    int_promoted_in_arithmetic => {
        body: "char a = 100; char b = 100;\nint c = a + b;\nprintf(\"%d\\n\", c);\nreturn 0;",
        expect: ["200"]
    },
    signed_unsigned_comparison => {
        // (unsigned)(-1) = 4294967295, so 1 > that is false → 1 (verified vs cc).
        body: "unsigned int u = 1;\nint s = -1;\nprintf(\"%d\\n\", u > (unsigned)s ? 0 : 1);\nreturn 0;",
        expect: ["1"]
    },
    float_double_promotion => {
        body: "float f = 1.0f / 3.0f;\ndouble d = 1.0 / 3.0;\nprintf(\"%d\\n\", d > f ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    zero_int_is_false => {
        body: "int x = 0;\nif (!x) printf(\"yes\\n\");\nreturn 0;",
        expect: ["yes"]
    },
    nonzero_int_is_true => {
        body: "int x = -1;\nif (x) printf(\"yes\\n\");\nreturn 0;",
        expect: ["yes"]
    }
}
