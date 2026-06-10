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
    puts_adds_newline => {
        body: "puts(\"hello\");\nreturn 0;",
        expect: ["hello"]
    },
    putchar_single_char => {
        body: "putchar('A');\nputchar('\\n');\nreturn 0;",
        expect: ["A"]
    },
    printf_no_newline_concatenates => {
        body: "printf(\"a\"); printf(\"b\"); printf(\"c\\n\");\nreturn 0;",
        expect: ["abc"]
    },
    printf_repeated_lines => {
        body: "for (int i = 1; i <= 3; i++) printf(\"%d\\n\", i);\nreturn 0;",
        expect: ["1", "2", "3"]
    },
    printf_table_format => {
        body: "printf(\"%5s %10s\\n\", \"key\", \"value\");\nprintf(\"%5d %10d\\n\", 1, 100);\nreturn 0;",
        expect: ["  key      value", "    1        100"]
    },
    printf_mixed_types_single_line => {
        body: "printf(\"%s=%d (%.2f)\\n\", \"pi\", 3, 3.14);\nreturn 0;",
        expect: ["pi=3 (3.14)"]
    },
    fputs_to_stdout => {
        body: "fputs(\"line\\n\", stdout);\nreturn 0;",
        expect: ["line"]
    },
    printf_long_integer_format => {
        body: "long x = 1234567890L;\nprintf(\"%ld\\n\", x);\nreturn 0;",
        expect: ["1234567890"]
    },
    printf_pointer_is_non_null => {
        body: "int x = 1;\nprintf(\"%d\\n\", &x != NULL ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    }
}
