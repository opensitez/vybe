use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<ctype.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    char_arithmetic_increment => {
        body: "char c = 'A';\nprintf(\"%c\\n\", c + 1);\nreturn 0;",
        expect: ["B"]
    },
    char_to_int_cast => {
        body: "char c = 'Z';\nprintf(\"%d\\n\", (int)c);\nreturn 0;",
        expect: ["90"]
    },
    int_to_char_cast => {
        body: "int n = 65;\nprintf(\"%c\\n\", (char)n);\nreturn 0;",
        expect: ["A"]
    },
    char_comparison => {
        body: "char a = 'a'; char b = 'b';\nprintf(\"%d\\n\", a < b ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    char_in_string_iteration => {
        body: r#"
char s[] = "abc";
for (int i = 0; s[i] != '\0'; i++) {
    printf("%c\n", s[i]);
}
return 0;
"#,
        expect: ["a", "b", "c"]
    },
    char_upper_to_lower => {
        body: "char c = 'A';\nprintf(\"%c\\n\", c + 32);\nreturn 0;",
        expect: ["a"]
    },
    char_lower_to_upper => {
        body: "char c = 'z';\nprintf(\"%c\\n\", c - 32);\nreturn 0;",
        expect: ["Z"]
    },
    char_escape_sequences => {
        body: "printf(\"%c%c%c\\n\", '\\t', '\\n', '\\\\');\nreturn 0;",
        expect: ["\t\n\\"]
    },
    char_null_terminator => {
        body: "char s[4] = {'a', 'b', 'c', '\\0'};\nprintf(\"%s\\n\", s);\nreturn 0;",
        expect: ["abc"]
    },
    char_array_as_string => {
        body: "char s[] = \"hello\";\nprintf(\"%d\\n\", (int)s[4]);\nreturn 0;",
        expect: ["111"]
    }
}
