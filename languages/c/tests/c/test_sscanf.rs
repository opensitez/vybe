use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<string.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    sscanf_integer => {
        body: "int n;\nsscanf(\"42\", \"%d\", &n);\nprintf(\"%d\\n\", n);\nreturn 0;",
        expect: ["42"]
    },
    sscanf_float => {
        body: "float f;\nsscanf(\"3.14\", \"%f\", &f);\nprintf(\"%.2f\\n\", f);\nreturn 0;",
        expect: ["3.14"]
    },
    sscanf_string => {
        body: "char buf[20];\nsscanf(\"hello\", \"%s\", buf);\nprintf(\"%s\\n\", buf);\nreturn 0;",
        expect: ["hello"]
    },
    sscanf_multiple_values => {
        body: "int a; float b;\nsscanf(\"10 2.5\", \"%d %f\", &a, &b);\nprintf(\"%d %.1f\\n\", a, b);\nreturn 0;",
        expect: ["10 2.5"]
    },
    sscanf_returns_count => {
        body: "int a, b;\nint n = sscanf(\"1 2\", \"%d %d\", &a, &b);\nprintf(\"%d\\n\", n);\nreturn 0;",
        expect: ["2"]
    },
    strtok_splits_on_delimiter => {
        body: r#"
char s[] = "a,b,c";
char *tok = strtok(s, ",");
while (tok != NULL) {
    printf("%s\n", tok);
    tok = strtok(NULL, ",");
}
return 0;
"#,
        expect: ["a", "b", "c"]
    },
    strtok_multiple_delimiters => {
        body: r#"
char s[] = "one two\tthree";
char *tok = strtok(s, " \t");
while (tok != NULL) {
    printf("%s\n", tok);
    tok = strtok(NULL, " \t");
}
return 0;
"#,
        expect: ["one", "two", "three"]
    },
    strtok_returns_null_when_done => {
        body: r#"
char s[] = "x";
char *tok = strtok(s, ",");
char *next = strtok(NULL, ",");
printf("%s %d\n", tok, next == NULL ? 1 : 0);
return 0;
"#,
        expect: ["x 1"]
    }
}
