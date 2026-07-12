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
    volatile_int_reads_correctly => {
        body: "volatile int x = 42;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["42"]
    },
    volatile_int_writes_and_reads => {
        body: "volatile int x = 0;\nx = 99;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["99"]
    },
    volatile_pointer => {
        body: "int val = 5;\nvolatile int *p = &val;\nprintf(\"%d\\n\", *p);\nreturn 0;",
        expect: ["5"]
    },
    volatile_in_loop => {
        body: "volatile int count = 0;\nfor (int i = 0; i < 5; i++) count++;\nprintf(\"%d\\n\", count);\nreturn 0;",
        expect: ["5"]
    }
}
