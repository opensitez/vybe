use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<stdlib.h>", "<string.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    int_to_string_sprintf => {
        body: "char buf[20];\nsprintf(buf, \"%d\", 12345);\nprintf(\"%s\\n\", buf);\nreturn 0;",
        expect: ["12345"]
    },
    float_to_string_sprintf => {
        body: "char buf[20];\nsprintf(buf, \"%.3f\", 3.14159);\nprintf(\"%s\\n\", buf);\nreturn 0;",
        expect: ["3.142"]
    },
    string_to_int_atoi => {
        body: "int n = atoi(\"-42\");\nprintf(\"%d\\n\", n);\nreturn 0;",
        expect: ["-42"]
    },
    string_to_float_atof => {
        body: "double d = atof(\"2.718\");\nprintf(\"%.3f\\n\", d);\nreturn 0;",
        expect: ["2.718"]
    },
    string_compare_equal => {
        body: "printf(\"%d\\n\", strcmp(\"abc\", \"abc\"));\nreturn 0;",
        expect: ["0"]
    },
    string_compare_less => {
        body: "printf(\"%d\\n\", strcmp(\"abc\", \"abd\") < 0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    string_compare_greater => {
        body: "printf(\"%d\\n\", strcmp(\"b\", \"a\") > 0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    string_concat_into_buffer => {
        body: "char dst[20] = \"hello\";\nstrcat(dst, \" world\");\nprintf(\"%s\\n\", dst);\nreturn 0;",
        expect: ["hello world"]
    },
    hex_int_to_string => {
        body: "char buf[10];\nsprintf(buf, \"%x\", 255);\nprintf(\"%s\\n\", buf);\nreturn 0;",
        expect: ["ff"]
    },
    number_formatted_with_width => {
        body: "char buf[10];\nsprintf(buf, \"%05d\", 42);\nprintf(\"%s\\n\", buf);\nreturn 0;",
        expect: ["00042"]
    }
}
