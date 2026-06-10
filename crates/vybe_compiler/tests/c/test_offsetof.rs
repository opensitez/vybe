use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<stddef.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    offsetof_first_field_is_zero => {
        declarations: "struct S { int a; int b; int c; };",
        body: "printf(\"%d\\n\", (int)offsetof(struct S, a));\nreturn 0;",
        expect: ["0"]
    },
    offsetof_second_field => {
        declarations: "struct S { int a; int b; };",
        body: "printf(\"%d\\n\", (int)offsetof(struct S, b));\nreturn 0;",
        expect: ["4"]
    },
    offsetof_third_field_aligned => {
        declarations: "struct S { int a; int b; int c; };",
        body: "printf(\"%d\\n\", (int)offsetof(struct S, c));\nreturn 0;",
        expect: ["8"]
    },
    offsetof_char_field => {
        declarations: "struct S { char a; char b; };",
        body: "printf(\"%d\\n\", (int)offsetof(struct S, b));\nreturn 0;",
        expect: ["1"]
    },
    offsetof_used_for_field_access => {
        declarations: "struct S { int x; int y; };",
        body: r#"
struct S s = {10, 20};
char *base = (char*)&s;
int *yp = (int*)(base + offsetof(struct S, y));
printf("%d\n", *yp);
return 0;
"#,
        expect: ["20"]
    }
}
