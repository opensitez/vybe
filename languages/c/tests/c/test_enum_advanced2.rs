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

c_cases! {
    enum_bitmask_flags => {
        declarations: "typedef enum { NONE=0, READ=1, WRITE=2, EXEC=4 } Perms;",
        body: "Perms p = READ | WRITE;\nprintf(\"%d\\n\", (p & READ) != 0 ? 1 : 0);\nprintf(\"%d\\n\", (p & EXEC) != 0 ? 1 : 0);\nreturn 0;",
        expect: ["1", "0"]
    },
    enum_iteration_via_loop => {
        declarations: "typedef enum { A=0, B=1, C=2, D=3 } Letter;\nconst char *names[] = {\"A\",\"B\",\"C\",\"D\"};",
        body: "for (Letter l = A; l <= D; l++) printf(\"%s\\n\", names[l]);\nreturn 0;",
        expect: ["A", "B", "C", "D"]
    },
    enum_in_switch_all_cases => {
        declarations: "typedef enum { X, Y, Z } Axis;",
        body: r#"
Axis a = Y;
switch (a) {
    case X: printf("x\n"); break;
    case Y: printf("y\n"); break;
    case Z: printf("z\n"); break;
}
return 0;
"#,
        expect: ["y"]
    },
    enum_negative_values => {
        declarations: "enum Err { ERR_OK=0, ERR_FAIL=-1, ERR_FATAL=-99 };",
        body: "printf(\"%d %d\\n\", ERR_FAIL, ERR_FATAL);\nreturn 0;",
        expect: ["-1 -99"]
    },
    enum_as_array_index => {
        declarations: "enum Color { RED=0, GREEN=1, BLUE=2 };\nconst char *color_names[] = {\"red\",\"green\",\"blue\"};",
        body: "printf(\"%s\\n\", color_names[GREEN]);\nreturn 0;",
        expect: ["green"]
    },
    enum_forward_and_typedef => {
        declarations: "typedef enum Status Status;\nenum Status { PENDING, RUNNING, DONE };",
        body: "Status s = RUNNING;\nprintf(\"%d\\n\", s);\nreturn 0;",
        expect: ["1"]
    }
}
