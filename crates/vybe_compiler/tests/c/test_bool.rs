use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<stdbool.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    bool_true_value => {
        body: "bool t = true;\nprintf(\"%d\\n\", t);\nreturn 0;",
        expect: ["1"]
    },
    bool_false_value => {
        body: "bool f = false;\nprintf(\"%d\\n\", f);\nreturn 0;",
        expect: ["0"]
    },
    bool_from_comparison => {
        body: "bool result = (5 > 3);\nprintf(\"%d\\n\", result);\nreturn 0;",
        expect: ["1"]
    },
    bool_negation => {
        body: "bool t = true;\nbool f = !t;\nprintf(\"%d %d\\n\", t, f);\nreturn 0;",
        expect: ["1 0"]
    },
    bool_in_if_condition => {
        body: "bool flag = true;\nif (flag) printf(\"yes\\n\"); else printf(\"no\\n\");\nreturn 0;",
        expect: ["yes"]
    },
    bool_integer_coercion => {
        body: "bool b = 42;\nprintf(\"%d\\n\", b);\nreturn 0;",
        expect: ["1"]
    },
    bool_zero_is_false => {
        body: "bool b = 0;\nprintf(\"%d\\n\", b);\nreturn 0;",
        expect: ["0"]
    },
    _Bool_without_header => {
        body: "printf(\"%d %d\\n\", (_Bool)1, (_Bool)0);\nreturn 0;",
        expect: ["1 0"]
    },
    bool_and_logic => {
        body: "bool a = true; bool b = false;\nprintf(\"%d %d\\n\", a && b, a || b);\nreturn 0;",
        expect: ["0 1"]
    }
}
