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
    int8_min_value => {
        body: "printf(\"%d\\n\", INT8_MIN);\nreturn 0;",
        expect: ["-128"]
    },
    int8_max_value => {
        body: "printf(\"%d\\n\", INT8_MAX);\nreturn 0;",
        expect: ["127"]
    },
    uint8_max_value => {
        body: "printf(\"%u\\n\", UINT8_MAX);\nreturn 0;",
        expect: ["255"]
    },
    int16_min_value => {
        body: "printf(\"%d\\n\", INT16_MIN);\nreturn 0;",
        expect: ["-32768"]
    },
    int16_max_value => {
        body: "printf(\"%d\\n\", INT16_MAX);\nreturn 0;",
        expect: ["32767"]
    },
    uint16_max_value => {
        body: "printf(\"%u\\n\", UINT16_MAX);\nreturn 0;",
        expect: ["65535"]
    },
    int32_min_value => {
        body: "printf(\"%d\\n\", INT32_MIN);\nreturn 0;",
        expect: ["-2147483648"]
    },
    int32_max_value => {
        body: "printf(\"%d\\n\", INT32_MAX);\nreturn 0;",
        expect: ["2147483647"]
    },
    uint32_max_value => {
        body: "printf(\"%u\\n\", UINT32_MAX);\nreturn 0;",
        expect: ["4294967295"]
    },
    size_max_positive => {
        body: "printf(\"%d\\n\", SIZE_MAX > 0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    intptr_min_negative => {
        body: "printf(\"%d\\n\", INTPTR_MIN < 0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    ptrdiff_max_positive => {
        body: "printf(\"%d\\n\", PTRDIFF_MAX > 0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    }
}
