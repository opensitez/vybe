use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<string.h>", "<locale.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    strcoll_equal_strings => {
        body: "printf(\"%d\\n\", strcoll(\"abc\", \"abc\") == 0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    strcoll_less_than => {
        body: "printf(\"%d\\n\", strcoll(\"abc\", \"abd\") < 0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    strxfrm_result_length => {
        body: r#"
char buf[64];
size_t n = strxfrm(buf, "hello", sizeof(buf));
printf("%d\n", (int)n == 5 ? 1 : 0);
return 0;
"#,
        expect: ["1"]
    },
    strxfrm_sortable_result => {
        body: r#"
char a[64], b[64];
strxfrm(a, "abc", sizeof(a));
strxfrm(b, "abd", sizeof(b));
printf("%d\n", strcmp(a, b) < 0 ? 1 : 0);
return 0;
"#,
        expect: ["1"]
    },
    setlocale_c_locale => {
        body: "char *loc = setlocale(LC_ALL, \"C\");\nprintf(\"%d\\n\", loc != NULL ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    }
}
