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
    fn_ptr_array_dispatch => {
        declarations: r#"
int add(int a, int b) { return a + b; }
int sub(int a, int b) { return a - b; }
int mul(int a, int b) { return a * b; }
typedef int (*BinOp)(int, int);
"#,
        body: r#"
BinOp ops[3] = {add, sub, mul};
printf("%d %d %d\n", ops[0](6,2), ops[1](6,2), ops[2](6,2));
return 0;
"#,
        expect: ["8 4 12"]
    },
    fn_ptr_assigned_and_called => {
        declarations: "int double_it(int x) { return x * 2; }\nint triple_it(int x) { return x * 3; }",
        body: r#"
int (*fn)(int);
fn = double_it;
printf("%d\n", fn(5));
fn = triple_it;
printf("%d\n", fn(5));
return 0;
"#,
        expect: ["10", "15"]
    },
    fn_ptr_as_parameter => {
        declarations: r#"
int apply_twice(int x, int (*f)(int)) { return f(f(x)); }
int increment(int x) { return x + 1; }
"#,
        body: "printf(\"%d\\n\", apply_twice(5, increment));\nreturn 0;",
        expect: ["7"]
    },
    fn_ptr_returned_from_function => {
        declarations: r#"
int add_n(int x) { return x + 10; }
int mul_n(int x) { return x * 10; }
typedef int (*Transform)(int);
Transform get_transform(int which) { return which == 0 ? add_n : mul_n; }
"#,
        body: r#"
int (*f)(int) = get_transform(1);
printf("%d\n", f(5));
return 0;
"#,
        expect: ["50"]
    },
    fn_ptr_null_check => {
        declarations: "int do_thing(void) { return 42; }",
        body: "int (*f)(void) = NULL;\nf = do_thing;\nprintf(\"%d\\n\", f != NULL ? f() : -1);\nreturn 0;",
        expect: ["42"]
    }
}
