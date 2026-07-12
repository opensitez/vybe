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
fn block_scope_shadows_outer() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int x = 1;\n{ int x = 2; printf(\"%d\\n\", x); }\nprintf(\"%d\\n\", x);\nreturn 0;",
        &["2", "1"],
    );
}

#[test]
fn nested_block_scopes() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int x = 10;\n{ int x = 20; { int x = 30; printf(\"%d\\n\", x); } printf(\"%d\\n\", x); }\nprintf(\"%d\\n\", x);\nreturn 0;",
        &["30", "20", "10"],
    );
}

#[test]
fn for_loop_variable_scoped() {
    assert_program(
        &["<stdio.h>"],
        "",
        "for (int i = 0; i < 2; i++) printf(\"%d\\n\", i);\nreturn 0;",
        &["0", "1"],
    );
}

#[test]
fn if_else_variable_in_block() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int x = 5;\nif (x > 0) { int y = x * 2; printf(\"%d\\n\", y); }\nreturn 0;",
        &["10"],
    );
}

#[test]
fn while_loop_scope() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int n = 3;\nwhile (n > 0) { int tmp = n * n; printf(\"%d\\n\", tmp); n--; }\nreturn 0;",
        &["9", "4", "1"],
    );
}

c_cases! {
    function_scope_isolation => {
        declarations: "int outer = 100;\nvoid fn() { int outer = 200; printf(\"%d\\n\", outer); }",
        body: "fn();\nprintf(\"%d\\n\", outer);\nreturn 0;",
        expect: ["200", "100"]
    }
}
