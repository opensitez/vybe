use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

#[test]
fn const_local_integer() {
    assert_program(
        &["<stdio.h>"],
        "",
        "const int x = 42;\nprintf(\"%d\\n\", x);\nreturn 0;",
        &["42"],
    );
}

#[test]
fn const_pointer_arithmetic() {
    assert_program(
        &["<stdio.h>"],
        "",
        "const int arr[] = {10,20,30};\nconst int *p = arr;\nprintf(\"%d %d\\n\", *p, *(p+2));\nreturn 0;",
        &["10 30"],
    );
}

#[test]
fn const_in_for_loop_bound() {
    assert_program(
        &["<stdio.h>"],
        "",
        "const int N = 4;\nint sum = 0;\nfor (int i = 1; i <= N; i++) sum += i;\nprintf(\"%d\\n\", sum);\nreturn 0;",
        &["10"],
    );
}

c_cases! {
    const_global_integer => {
        declarations: "const int LIMIT = 100;",
        body: "printf(\"%d\\n\", LIMIT);\nreturn 0;",
        expect: ["100"]
    },
    const_string_literal => {
        declarations: "const char *msg = \"hello\";",
        body: "printf(\"%s\\n\", msg);\nreturn 0;",
        expect: ["hello"]
    },
    const_function_parameter => {
        declarations: "int strlen_safe(const char *s) { int n = 0; while (*s++) n++; return n; }",
        body: "printf(\"%d\\n\", strlen_safe(\"hello\"));\nreturn 0;",
        expect: ["5"]
    },
    const_struct_field => {
        declarations: "struct Rect { const int w; const int h; };",
        body: "struct Rect r = {3, 4};\nprintf(\"%d\\n\", r.w * r.h);\nreturn 0;",
        expect: ["12"]
    }
}
