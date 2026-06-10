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
    inline_function_basic => {
        declarations: "static inline int square(int x) { return x * x; }",
        body: "printf(\"%d\\n\", square(5));\nreturn 0;",
        expect: ["25"]
    },
    inline_function_called_multiple_times => {
        declarations: "static inline int add_one(int x) { return x + 1; }",
        body: "printf(\"%d %d %d\\n\", add_one(1), add_one(10), add_one(100));\nreturn 0;",
        expect: ["2 11 101"]
    },
    inline_function_with_condition => {
        declarations: "static inline int clamp(int x, int lo, int hi) { return x < lo ? lo : (x > hi ? hi : x); }",
        body: "printf(\"%d %d %d\\n\", clamp(-5, 0, 10), clamp(5, 0, 10), clamp(15, 0, 10));\nreturn 0;",
        expect: ["0 5 10"]
    },
    inline_function_returns_pointer => {
        declarations: "static inline const char* greet(int n) { return n > 0 ? \"positive\" : \"nonpositive\"; }",
        body: "printf(\"%s %s\\n\", greet(5), greet(-3));\nreturn 0;",
        expect: ["positive nonpositive"]
    },
    static_inline_avoids_symbol_conflict => {
        declarations: "static inline double halve(double x) { return x / 2.0; }",
        body: "printf(\"%.1f\\n\", halve(7.0));\nreturn 0;",
        expect: ["3.5"]
    }
}
