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
fn auto_keyword_explicit() {
    assert_program(
        &["<stdio.h>"],
        "",
        "auto int x = 42;\nprintf(\"%d\\n\", x);\nreturn 0;",
        &["42"],
    );
}

#[test]
fn register_keyword_integer() {
    assert_program(
        &["<stdio.h>"],
        "",
        "register int i;\nfor (i = 0; i < 3; i++) { printf(\"%d\\n\", i); }\nreturn 0;",
        &["0", "1", "2"],
    );
}

#[test]
fn register_keyword_char() {
    assert_program(
        &["<stdio.h>"],
        "",
        "register char c = 'A';\nprintf(\"%c\\n\", c);\nreturn 0;",
        &["A"],
    );
}

c_cases! {
    static_local_preserves_value => {
        declarations: "void counter() { static int n = 0; n++; printf(\"%d\\n\", n); }",
        body: "counter(); counter(); counter(); return 0;",
        expect: ["1", "2", "3"]
    },
    static_local_init_once => {
        declarations: "void init_once() { static int x = 100; x += 10; printf(\"%d\\n\", x); }",
        body: "init_once(); init_once(); return 0;",
        expect: ["110", "120"]
    },
    static_function_scope => {
        declarations: "static int double_it(int x) { return x * 2; }",
        body: "printf(\"%d\\n\", double_it(5)); return 0;",
        expect: ["10"]
    },
    extern_declaration => {
        declarations: "int global_val = 99;\nextern int global_val;",
        body: "printf(\"%d\\n\", global_val); return 0;",
        expect: ["99"]
    }
}
