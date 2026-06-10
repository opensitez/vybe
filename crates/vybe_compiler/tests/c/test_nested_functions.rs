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

// Forward declarations and mutual recursion patterns
c_cases! {
    forward_declaration_before_definition => {
        declarations: "int bar(int x);\nint foo(int x) { return bar(x + 1); }\nint bar(int x) { return x * 2; }",
        body: "printf(\"%d\\n\", foo(3));\nreturn 0;",
        expect: ["8"]
    },
    function_returning_struct => {
        declarations: "typedef struct { int x; int y; } Point;\nPoint make_point(int x, int y) { Point p = {x, y}; return p; }",
        body: "Point p = make_point(3, 4);\nprintf(\"%d %d\\n\", p.x, p.y);\nreturn 0;",
        expect: ["3 4"]
    },
    function_taking_array_and_size => {
        declarations: "int array_sum(int *arr, int n) { int s = 0; for (int i = 0; i < n; i++) s += arr[i]; return s; }",
        body: "int data[] = {1,2,3,4,5};\nprintf(\"%d\\n\", array_sum(data, 5));\nreturn 0;",
        expect: ["15"]
    },
    function_modifying_via_pointer => {
        declarations: "void double_val(int *p) { *p *= 2; }",
        body: "int x = 7;\ndouble_val(&x);\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["14"]
    },
    callback_via_function_pointer => {
        declarations: "int apply(int x, int (*fn)(int)) { return fn(x); }\nint triple(int x) { return x * 3; }",
        body: "printf(\"%d\\n\", apply(5, triple));\nreturn 0;",
        expect: ["15"]
    },
    function_with_multiple_return_paths => {
        declarations: "const char* classify(int n) { if (n < 0) return \"neg\"; if (n == 0) return \"zero\"; return \"pos\"; }",
        body: "printf(\"%s %s %s\\n\", classify(-1), classify(0), classify(1));\nreturn 0;",
        expect: ["neg zero pos"]
    }
}
