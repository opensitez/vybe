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
    hex_literal_uppercase => {
        body: "printf(\"%d\\n\", 0xFF);\nreturn 0;",
        expect: ["255"]
    },
    hex_literal_lowercase => {
        body: "printf(\"%d\\n\", 0xab);\nreturn 0;",
        expect: ["171"]
    },
    octal_literal => {
        body: "printf(\"%d\\n\", 0755);\nreturn 0;",
        expect: ["493"]
    },
    octal_zero => {
        body: "printf(\"%d\\n\", 00);\nreturn 0;",
        expect: ["0"]
    },
    binary_literal_zero => {
        body: "printf(\"%d\\n\", 0b0);\nreturn 0;",
        expect: ["0"]
    },
    binary_literal_one => {
        body: "printf(\"%d\\n\", 0b1);\nreturn 0;",
        expect: ["1"]
    },
    binary_literal_byte => {
        body: "printf(\"%d\\n\", 0b11001010);\nreturn 0;",
        expect: ["202"]
    },
    float_scientific_notation => {
        body: "printf(\"%.0f\\n\", 1.5e3);\nreturn 0;",
        expect: ["1500"]
    },
    float_negative_exponent => {
        body: "printf(\"%.4f\\n\", 1.0e-3);\nreturn 0;",
        expect: ["0.0010"]
    },
    long_literal_suffix => {
        body: "printf(\"%ld\\n\", 2147483648L);\nreturn 0;",
        expect: ["2147483648"]
    },
    unsigned_literal_suffix => {
        body: "printf(\"%u\\n\", 4294967295U);\nreturn 0;",
        expect: ["4294967295"]
    },
    char_literal_escape_hex => {
        body: "printf(\"%d\\n\", '\\x41');\nreturn 0;",
        expect: ["65"]
    },
    char_literal_escape_octal => {
        body: "printf(\"%d\\n\", '\\101');\nreturn 0;",
        expect: ["65"]
    }
}
