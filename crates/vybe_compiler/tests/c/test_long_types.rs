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
    long_long_basic_value => {
        body: "long long x = 1000000000LL * 2;\nprintf(\"%lld\\n\", x);\nreturn 0;",
        expect: ["2000000000"]
    },
    long_long_max_overflow_detection => {
        body: "long long a = 9000000000LL;\nlong long b = 1000000000LL;\nprintf(\"%lld\\n\", a + b);\nreturn 0;",
        expect: ["10000000000"]
    },
    unsigned_long_long_basic => {
        body: "unsigned long long x = 18000000000ULL;\nprintf(\"%llu\\n\", x);\nreturn 0;",
        expect: ["18000000000"]
    },
    long_double_basic_value => {
        body: "long double x = 3.14159265358979L;\nprintf(\"%.5Lf\\n\", x);\nreturn 0;",
        expect: ["3.14159"]
    },
    long_int_arithmetic => {
        body: "long int a = 2147483647L;\nlong int b = 1L;\nprintf(\"%ld\\n\", a + b);\nreturn 0;",
        expect: ["2147483648"]
    },
    unsigned_long_basic => {
        body: "unsigned long x = 4000000000UL;\nprintf(\"%lu\\n\", x);\nreturn 0;",
        expect: ["4000000000"]
    },
    long_long_multiplication => {
        body: "long long a = 100000LL;\nlong long b = 100000LL;\nprintf(\"%lld\\n\", a * b);\nreturn 0;",
        expect: ["10000000000"]
    },
    short_int_basic => {
        body: "short int x = 32767;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["32767"]
    },
    unsigned_short_basic => {
        body: "unsigned short x = 65535;\nprintf(\"%u\\n\", x);\nreturn 0;",
        expect: ["65535"]
    },
    signed_char_range => {
        body: "signed char c = -128;\nprintf(\"%d\\n\", c);\nreturn 0;",
        expect: ["-128"]
    },
    unsigned_char_range => {
        body: "unsigned char c = 255;\nprintf(\"%u\\n\", c);\nreturn 0;",
        expect: ["255"]
    }
}
