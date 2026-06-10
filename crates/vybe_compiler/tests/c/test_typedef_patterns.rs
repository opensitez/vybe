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
    typedef_unsigned_byte => {
        declarations: "typedef unsigned char u8;",
        body: "u8 x = 200;\nprintf(\"%u\\n\", x);\nreturn 0;",
        expect: ["200"]
    },
    typedef_signed_short => {
        declarations: "typedef short int16_t_compat;",
        body: "int16_t_compat x = -1000;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["-1000"]
    },
    typedef_function_signature => {
        declarations: "typedef void (*Handler)(int);\nvoid on_event(int code) { printf(\"%d\\n\", code); }",
        body: "Handler h = on_event;\nh(42);\nreturn 0;",
        expect: ["42"]
    },
    typedef_struct_anonymous => {
        declarations: "typedef struct { float x; float y; } Vector2;",
        body: "Vector2 v = {3.0f, 4.0f};\nprintf(\"%.0f %.0f\\n\", v.x, v.y);\nreturn 0;",
        expect: ["3 4"]
    },
    typedef_nested_struct => {
        declarations: "typedef struct { int r; int g; int b; } Color;\ntypedef struct { Color fg; Color bg; } Theme;",
        body: "Theme t = {{255,0,0},{0,0,255}};\nprintf(\"%d %d\\n\", t.fg.r, t.bg.b);\nreturn 0;",
        expect: ["255 255"]
    },
    typedef_pointer_opaque => {
        declarations: "struct _Handle { int id; };\ntypedef struct _Handle *Handle;",
        body: "struct _Handle h = {99};\nHandle p = &h;\nprintf(\"%d\\n\", p->id);\nreturn 0;",
        expect: ["99"]
    },
    typedef_multidim_array => {
        declarations: "typedef int Matrix3x3[3][3];",
        body: "Matrix3x3 m = {{1,0,0},{0,1,0},{0,0,1}};\nprintf(\"%d %d %d\\n\", m[0][0], m[1][1], m[2][2]);\nreturn 0;",
        expect: ["1 1 1"]
    }
}
