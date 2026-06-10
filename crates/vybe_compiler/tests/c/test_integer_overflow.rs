use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<limits.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    unsigned_overflow_wraps_to_zero => {
        body: "unsigned int x = UINT_MAX;\nx++;\nprintf(\"%u\\n\", x);\nreturn 0;",
        expect: ["0"]
    },
    unsigned_underflow_wraps_to_max => {
        body: "unsigned int x = 0;\nx--;\nprintf(\"%u\\n\", x);\nreturn 0;",
        expect: ["4294967295"]
    },
    unsigned_addition_wraps => {
        body: "unsigned char x = 255;\nunsigned char y = x + 1;\nprintf(\"%u\\n\", y);\nreturn 0;",
        expect: ["0"]
    },
    char_overflow_wraps => {
        body: "unsigned char c = 250;\nc += 10;\nprintf(\"%u\\n\", c);\nreturn 0;",
        expect: ["4"]
    },
    int_modulo_positive => {
        body: "printf(\"%d\\n\", 17 % 5);\nreturn 0;",
        expect: ["2"]
    },
    int_division_truncates => {
        body: "printf(\"%d %d\\n\", 7 / 2, -7 / 2);\nreturn 0;",
        expect: ["3 -3"]
    },
    int_modulo_negative => {
        body: "printf(\"%d\\n\", -7 % 3);\nreturn 0;",
        expect: ["-1"]
    },
    bit_shift_left => {
        body: "printf(\"%d\\n\", 1 << 8);\nreturn 0;",
        expect: ["256"]
    },
    bit_shift_right => {
        body: "printf(\"%d\\n\", 256 >> 3);\nreturn 0;",
        expect: ["32"]
    },
    unsigned_shift_right_logical => {
        body: "unsigned int x = 0x80000000U;\nprintf(\"%u\\n\", x >> 1);\nreturn 0;",
        expect: ["1073741824"]
    }
}
