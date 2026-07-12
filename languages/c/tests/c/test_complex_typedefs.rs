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
    typedef_function_pointer => {
        declarations: "typedef int (*BinaryOp)(int, int);\nint add(int a, int b) { return a + b; }\nint mul(int a, int b) { return a * b; }",
        body: "BinaryOp ops[2] = {add, mul};\nprintf(\"%d %d\\n\", ops[0](3,4), ops[1](3,4));\nreturn 0;",
        expect: ["7 12"]
    },
    typedef_pointer_to_struct => {
        declarations: "struct Node { int val; };\ntypedef struct Node *NodePtr;",
        body: "struct Node n; n.val = 55;\nNodePtr p = &n;\nprintf(\"%d\\n\", p->val);\nreturn 0;",
        expect: ["55"]
    },
    typedef_array_type => {
        declarations: "typedef int IntArr[4];",
        body: "IntArr a = {1, 2, 3, 4};\nprintf(\"%d %d\\n\", a[0], a[3]);\nreturn 0;",
        expect: ["1 4"]
    },
    typedef_enum => {
        declarations: "typedef enum { RED, GREEN, BLUE } Color;",
        body: "Color c = GREEN;\nprintf(\"%d\\n\", c);\nreturn 0;",
        expect: ["1"]
    },
    typedef_struct_self_named => {
        declarations: "typedef struct Vec2 { float x; float y; } Vec2;",
        body: "Vec2 v = {1.0f, 2.0f};\nprintf(\"%.0f %.0f\\n\", v.x, v.y);\nreturn 0;",
        expect: ["1 2"]
    },
    typedef_chained => {
        declarations: "typedef int Int;\ntypedef Int MyInt;\ntypedef MyInt FinalInt;",
        body: "FinalInt x = 42;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["42"]
    },
    typedef_const_pointer => {
        declarations: "typedef const char *CStr;",
        body: "CStr s = \"hello\";\nprintf(\"%s\\n\", s);\nreturn 0;",
        expect: ["hello"]
    }
}
