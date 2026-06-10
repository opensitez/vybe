use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<stdlib.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<stdlib.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    pointer_to_pointer_basic => {
        body: "int x = 42;\nint *p = &x;\nint **pp = &p;\nprintf(\"%d\\n\", **pp);\nreturn 0;",
        expect: ["42"]
    },
    pointer_to_pointer_write_through => {
        body: "int x = 0;\nint *p = &x;\nint **pp = &p;\n**pp = 99;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["99"]
    },
    void_pointer_cast => {
        body: "int x = 55;\nvoid *vp = &x;\nint *ip = (int*)vp;\nprintf(\"%d\\n\", *ip);\nreturn 0;",
        expect: ["55"]
    },
    const_pointer_to_data => {
        body: "int x = 10; int y = 20;\nconst int *p = &x;\np = &y;\nprintf(\"%d\\n\", *p);\nreturn 0;",
        expect: ["20"]
    },
    pointer_to_const_data => {
        body: "int x = 10;\nint * const p = &x;\n*p = 99;\nprintf(\"%d\\n\", *p);\nreturn 0;",
        expect: ["99"]
    },
    null_pointer_check => {
        body: "int *p = NULL;\nprintf(\"%d\\n\", p == NULL ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    pointer_difference => {
        body: "int arr[5];\nint *a = &arr[1];\nint *b = &arr[4];\nprintf(\"%d\\n\", (int)(b - a));\nreturn 0;",
        expect: ["3"]
    },
    dynamic_array_via_malloc => {
        body: r#"
int n = 4;
int *arr = (int*)malloc(n * sizeof(int));
for (int i = 0; i < n; i++) arr[i] = i * i;
printf("%d %d %d %d\n", arr[0], arr[1], arr[2], arr[3]);
free(arr);
return 0;
"#,
        expect: ["0 1 4 9"]
    },
    pointer_comparison => {
        body: "int arr[3];\nint *p = arr;\nint *q = arr + 2;\nprintf(\"%d\\n\", q > p ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    }
}
