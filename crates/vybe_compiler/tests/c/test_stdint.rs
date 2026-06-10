use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<stdint.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    int8_t_range => {
        body: "int8_t x = 127;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["127"]
    },
    int8_t_negative => {
        body: "int8_t x = -128;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["-128"]
    },
    uint8_t_max => {
        body: "uint8_t x = 255;\nprintf(\"%u\\n\", x);\nreturn 0;",
        expect: ["255"]
    },
    int16_t_range => {
        body: "int16_t x = 32767;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["32767"]
    },
    uint16_t_max => {
        body: "uint16_t x = 65535;\nprintf(\"%u\\n\", x);\nreturn 0;",
        expect: ["65535"]
    },
    int32_t_range => {
        body: "int32_t x = 2147483647;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["2147483647"]
    },
    uint32_t_max => {
        body: "uint32_t x = 4294967295U;\nprintf(\"%u\\n\", x);\nreturn 0;",
        expect: ["4294967295"]
    },
    int64_t_range => {
        body: "int64_t x = 9000000000LL;\nprintf(\"%lld\\n\", (long long)x);\nreturn 0;",
        expect: ["9000000000"]
    },
    uint64_t_max => {
        body: "uint64_t x = 18000000000ULL;\nprintf(\"%llu\\n\", (unsigned long long)x);\nreturn 0;",
        expect: ["18000000000"]
    },
    int_least8_t_basic => {
        body: "int_least8_t x = 100;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["100"]
    },
    int_fast32_t_arithmetic => {
        body: "int_fast32_t a = 100;\nint_fast32_t b = 200;\nprintf(\"%d\\n\", (int)(a + b));\nreturn 0;",
        expect: ["300"]
    },
    intptr_t_holds_address => {
        body: "int x = 42;\nintptr_t p = (intptr_t)&x;\nprintf(\"%d\\n\", *((int*)p));\nreturn 0;",
        expect: ["42"]
    }
}
