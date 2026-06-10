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
    if_else_chain => {
        body: r#"
int x = 5;
if (x < 0) printf("neg\n");
else if (x == 0) printf("zero\n");
else if (x < 10) printf("small\n");
else printf("large\n");
return 0;
"#,
        expect: ["small"]
    },
    nested_if_else => {
        body: r#"
int a = 1, b = 2;
if (a > 0) {
    if (b > 0) printf("both pos\n");
    else printf("a pos\n");
} else {
    printf("a neg\n");
}
return 0;
"#,
        expect: ["both pos"]
    },
    dangling_else_attaches_to_nearest => {
        body: "int x = 1;\nif (x > 0) if (x > 10) printf(\"big\\n\"); else printf(\"small\\n\");\nreturn 0;",
        expect: ["small"]
    },
    ternary_nested => {
        body: "int x = 5;\nconst char *s = x < 0 ? \"neg\" : (x == 0 ? \"zero\" : \"pos\");\nprintf(\"%s\\n\", s);\nreturn 0;",
        expect: ["pos"]
    },
    logical_and_short_circuits => {
        body: "int x = 0;\nint y = x && (1/x > 0);\nprintf(\"%d\\n\", y);\nreturn 0;",
        expect: ["0"]
    },
    logical_or_short_circuits => {
        body: "int x = 1;\nint y = x || (1/0 > 0);\nprintf(\"%d\\n\", y);\nreturn 0;",
        expect: ["1"]
    },
    comparison_chain_via_variables => {
        body: "int a=1, b=2, c=3;\nprintf(\"%d\\n\", a < b && b < c ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    null_pointer_in_condition => {
        body: "int *p = NULL;\nprintf(\"%d\\n\", p ? 1 : 0);\nreturn 0;",
        expect: ["0"]
    },
    string_literal_is_truthy => {
        body: "const char *s = \"hello\";\nprintf(\"%d\\n\", s ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    }
}
