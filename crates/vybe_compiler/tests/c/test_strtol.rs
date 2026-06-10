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
    strtol_decimal => {
        body: "char *end;\nlong v = strtol(\"42\", &end, 10);\nprintf(\"%ld\\n\", v);\nreturn 0;",
        expect: ["42"]
    },
    strtol_negative => {
        body: "char *end;\nlong v = strtol(\"-99\", &end, 10);\nprintf(\"%ld\\n\", v);\nreturn 0;",
        expect: ["-99"]
    },
    strtol_hex_base => {
        body: "char *end;\nlong v = strtol(\"0xff\", &end, 16);\nprintf(\"%ld\\n\", v);\nreturn 0;",
        expect: ["255"]
    },
    strtol_octal_base => {
        body: "char *end;\nlong v = strtol(\"010\", &end, 8);\nprintf(\"%ld\\n\", v);\nreturn 0;",
        expect: ["8"]
    },
    strtol_auto_base => {
        body: "char *end;\nlong v = strtol(\"0x1A\", &end, 0);\nprintf(\"%ld\\n\", v);\nreturn 0;",
        expect: ["26"]
    },
    strtol_endptr_advances => {
        body: "char *end;\nchar s[] = \"123abc\";\nlong v = strtol(s, &end, 10);\nprintf(\"%ld %s\\n\", v, end);\nreturn 0;",
        expect: ["123 abc"]
    },
    strtoul_basic => {
        body: "char *end;\nunsigned long v = strtoul(\"4294967295\", &end, 10);\nprintf(\"%lu\\n\", v);\nreturn 0;",
        expect: ["4294967295"]
    },
    strtod_decimal_point => {
        body: "char *end;\ndouble v = strtod(\"3.14\", &end);\nprintf(\"%.2f\\n\", v);\nreturn 0;",
        expect: ["3.14"]
    },
    strtod_scientific => {
        body: "char *end;\ndouble v = strtod(\"1.5e2\", &end);\nprintf(\"%.1f\\n\", v);\nreturn 0;",
        expect: ["150.0"]
    },
    strtod_endptr_at_non_numeric => {
        body: "char *end;\nchar s[] = \"2.5xyz\";\nstrtod(s, &end);\nprintf(\"%s\\n\", end);\nreturn 0;",
        expect: ["xyz"]
    }
}
