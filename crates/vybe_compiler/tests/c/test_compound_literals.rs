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
fn compound_literal_array_assign() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int *p = (int[]){10, 20, 30};\nprintf(\"%d %d %d\\n\", p[0], p[1], p[2]);\nreturn 0;",
        &["10 20 30"],
    );
}

#[test]
fn compound_literal_array_in_expression() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int x = ((int[]){5, 6, 7})[1];\nprintf(\"%d\\n\", x);\nreturn 0;",
        &["6"],
    );
}

c_cases! {
    compound_literal_struct_direct_use => {
        declarations: "struct Point { int x; int y; };",
        body: "struct Point p = (struct Point){.x=3, .y=4};\nprintf(\"%d %d\\n\", p.x, p.y);\nreturn 0;",
        expect: ["3 4"]
    },
    compound_literal_struct_passed_to_fn => {
        declarations: "struct Point { int x; int y; };\nint sum_coords(struct Point p) { return p.x + p.y; }",
        body: "printf(\"%d\\n\", sum_coords((struct Point){.x=7, .y=8}));\nreturn 0;",
        expect: ["15"]
    },
    compound_literal_struct_field_access => {
        declarations: "struct Val { int n; };",
        body: "int v = ((struct Val){.n=99}).n;\nprintf(\"%d\\n\", v);\nreturn 0;",
        expect: ["99"]
    }
}
