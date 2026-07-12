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
}

c_cases! {
    comma_operator_in_for => {
        body: "int a=0, b=0;\nfor (a=1, b=10; a < 3; a++, b--) { printf(\"%d %d\\n\", a, b); }\nreturn 0;",
        expect: ["1 10", "2 9"]
    },
    conditional_expression_nesting => {
        body: "int x = 5;\nint y = x > 0 ? (x > 10 ? 10 : x) : 0;\nprintf(\"%d\\n\", y);\nreturn 0;",
        expect: ["5"]
    },
    sizeof_expression => {
        body: "int arr[10];\nprintf(\"%d\\n\", (int)(sizeof(arr) / sizeof(arr[0])));\nreturn 0;",
        expect: ["10"]
    },
    compound_assignment_chain => {
        body: "int x = 10;\nx += 5;\nx *= 2;\nx -= 3;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["27"]
    },
    bitwise_ops_combined => {
        // (0xAB & 0x0F) | (0xAB >> 4) = 0xB | 0xA = 11 (verified vs native cc).
        body: "int x = 0xAB;\nint y = (x & 0x0F) | (x >> 4);\nprintf(\"%d\\n\", y);\nreturn 0;",
        expect: ["11"]
    },
    short_circuit_and_no_side_effect => {
        body: "int x = 0;\nint y = 0;\n(y = 1) && (x = 2);\nprintf(\"%d %d\\n\", x, y);\nreturn 0;",
        expect: ["2 1"]
    },
    short_circuit_or_no_side_effect => {
        body: "int x = 0;\nint y = 1;\n(y = 0) || (x = 5);\nprintf(\"%d %d\\n\", x, y);\nreturn 0;",
        expect: ["5 0"]
    },
    cast_in_arithmetic => {
        body: "int a = 7, b = 2;\nfloat result = (float)a / b;\nprintf(\"%.1f\\n\", result);\nreturn 0;",
        expect: ["3.5"]
    },
    address_of_dereferenced => {
        body: "int x = 42;\nint *p = &x;\nprintf(\"%d\\n\", *(&(*p)));\nreturn 0;",
        expect: ["42"]
    },
    sequence_point_comma => {
        body: "int a = (1, 2, 3);\nprintf(\"%d\\n\", a);\nreturn 0;",
        expect: ["3"]
    }
}
