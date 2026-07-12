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
    bitfield_basic_read_write => {
        declarations: "struct Flags { unsigned int a : 1; unsigned int b : 1; unsigned int c : 1; };",
        body: "struct Flags f;\nf.a = 1; f.b = 0; f.c = 1;\nprintf(\"%u %u %u\\n\", f.a, f.b, f.c);\nreturn 0;",
        expect: ["1 0 1"]
    },
    bitfield_multi_bit_value => {
        declarations: "struct Packed { unsigned int low : 4; unsigned int high : 4; };",
        body: "struct Packed p;\np.low = 5; p.high = 12;\nprintf(\"%u %u\\n\", p.low, p.high);\nreturn 0;",
        expect: ["5 12"]
    },
    bitfield_overflow_wraps => {
        declarations: "struct B { unsigned int val : 2; };",
        body: "struct B b;\nb.val = 7;\nprintf(\"%u\\n\", b.val);\nreturn 0;",
        expect: ["3"]
    },
    bitfield_signed_value => {
        declarations: "struct S { int val : 4; };",
        body: "struct S s;\ns.val = -1;\nprintf(\"%d\\n\", s.val);\nreturn 0;",
        expect: ["-1"]
    },
    bitfield_initializer => {
        declarations: "struct Rgb { unsigned int r : 8; unsigned int g : 8; unsigned int b : 8; };",
        body: "struct Rgb color = {255, 128, 0};\nprintf(\"%u %u %u\\n\", color.r, color.g, color.b);\nreturn 0;",
        expect: ["255 128 0"]
    },
    bitfield_combined_with_regular_field => {
        declarations: "struct Mixed { int id; unsigned int flag : 1; unsigned int count : 7; };",
        body: "struct Mixed m;\nm.id = 42; m.flag = 1; m.count = 100;\nprintf(\"%d %u %u\\n\", m.id, m.flag, m.count);\nreturn 0;",
        expect: ["42 1 100"]
    }
}
