use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<stdlib.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    abs_positive_unchanged => {
        body: "printf(\"%d\\n\", abs(42));\nreturn 0;",
        expect: ["42"]
    },
    abs_negative_becomes_positive => {
        body: "printf(\"%d\\n\", abs(-42));\nreturn 0;",
        expect: ["42"]
    },
    labs_large_negative => {
        body: "printf(\"%ld\\n\", labs(-1000000L));\nreturn 0;",
        expect: ["1000000"]
    },
    div_quotient_and_remainder => {
        body: "div_t r = div(17, 5);\nprintf(\"%d %d\\n\", r.quot, r.rem);\nreturn 0;",
        expect: ["3 2"]
    },
    ldiv_quotient_and_remainder => {
        body: "ldiv_t r = ldiv(100L, 7L);\nprintf(\"%ld %ld\\n\", r.quot, r.rem);\nreturn 0;",
        expect: ["14 2"]
    },
    rand_returns_in_range => {
        body: "srand(42);\nint r = rand() % 10;\nprintf(\"%d\\n\", r >= 0 && r < 10 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    rand_reproducible_with_seed => {
        body: "srand(12345);\nint a = rand();\nsrand(12345);\nint b = rand();\nprintf(\"%d\\n\", a == b ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    atoi_basic => {
        body: "printf(\"%d\\n\", atoi(\"123\"));\nreturn 0;",
        expect: ["123"]
    },
    atof_basic => {
        body: "printf(\"%.2f\\n\", atof(\"3.14\"));\nreturn 0;",
        expect: ["3.14"]
    },
    atol_basic => {
        body: "printf(\"%ld\\n\", atol(\"9876\"));\nreturn 0;",
        expect: ["9876"]
    }
}
