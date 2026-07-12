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
fn array_designated_init_by_index() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int arr[5] = {[0]=10, [2]=20, [4]=30};\nprintf(\"%d %d %d %d %d\\n\", arr[0], arr[1], arr[2], arr[3], arr[4]);\nreturn 0;",
        &["10 0 20 0 30"],
    );
}

#[test]
fn array_designated_init_partial_zero_fill() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int arr[4] = {[1]=99};\nprintf(\"%d %d %d %d\\n\", arr[0], arr[1], arr[2], arr[3]);\nreturn 0;",
        &["0 99 0 0"],
    );
}

#[test]
fn array_designated_init_out_of_order() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int arr[3] = {[2]=3, [0]=1, [1]=2};\nprintf(\"%d %d %d\\n\", arr[0], arr[1], arr[2]);\nreturn 0;",
        &["1 2 3"],
    );
}

c_cases! {
    struct_designated_init_named_fields => {
        declarations: "struct Point { int x; int y; int z; };",
        body: "struct Point p = {.x=1, .z=3};\nprintf(\"%d %d %d\\n\", p.x, p.y, p.z);\nreturn 0;",
        expect: ["1 0 3"]
    },
    struct_designated_init_all_fields => {
        declarations: "struct Color { int r; int g; int b; };",
        body: "struct Color c = {.b=255, .g=128, .r=64};\nprintf(\"%d %d %d\\n\", c.r, c.g, c.b);\nreturn 0;",
        expect: ["64 128 255"]
    },
    struct_designated_init_nested => {
        declarations: "struct Inner { int a; int b; }; struct Outer { struct Inner in; int c; };",
        body: "struct Outer o = {.in={.a=1, .b=2}, .c=3};\nprintf(\"%d %d %d\\n\", o.in.a, o.in.b, o.c);\nreturn 0;",
        expect: ["1 2 3"]
    }
}
