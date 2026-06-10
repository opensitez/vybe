use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<stdlib.h>", "<string.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

#[test]
fn malloc_and_free_int() {
    assert_program(
        &["<stdio.h>", "<stdlib.h>"],
        "",
        "int *p = (int*)malloc(sizeof(int));\n*p = 99;\nprintf(\"%d\\n\", *p);\nfree(p);\nreturn 0;",
        &["99"],
    );
}

#[test]
fn calloc_zeroes_memory() {
    assert_program(
        &["<stdio.h>", "<stdlib.h>"],
        "",
        "int *p = (int*)calloc(3, sizeof(int));\nprintf(\"%d %d %d\\n\", p[0], p[1], p[2]);\nfree(p);\nreturn 0;",
        &["0 0 0"],
    );
}

#[test]
fn malloc_array_access() {
    assert_program(
        &["<stdio.h>", "<stdlib.h>"],
        "",
        r#"
int n = 5;
int *arr = (int*)malloc(n * sizeof(int));
for (int i = 0; i < n; i++) arr[i] = i * i;
printf("%d %d %d\n", arr[0], arr[2], arr[4]);
free(arr);
return 0;
"#,
        &["0 4 16"],
    );
}

#[test]
fn malloc_string_copy() {
    assert_program(
        &["<stdio.h>", "<stdlib.h>", "<string.h>"],
        "",
        r#"
const char *src = "hello";
char *dst = (char*)malloc(6);
strcpy(dst, src);
printf("%s\n", dst);
free(dst);
return 0;
"#,
        &["hello"],
    );
}

#[test]
fn realloc_preserves_data() {
    assert_program(
        &["<stdio.h>", "<stdlib.h>"],
        "",
        r#"
int *p = (int*)malloc(3 * sizeof(int));
p[0] = 1; p[1] = 2; p[2] = 3;
p = (int*)realloc(p, 5 * sizeof(int));
p[3] = 4; p[4] = 5;
printf("%d %d %d %d %d\n", p[0], p[1], p[2], p[3], p[4]);
free(p);
return 0;
"#,
        &["1 2 3 4 5"],
    );
}

c_cases! {
    calloc_struct_array => {
        declarations: "struct Point { int x; int y; };",
        body: r#"
struct Point *pts = (struct Point*)calloc(3, sizeof(struct Point));
pts[1].x = 5; pts[1].y = 10;
printf("%d %d %d %d\n", pts[0].x, pts[0].y, pts[1].x, pts[1].y);
free(pts);
return 0;
"#,
        expect: ["0 0 5 10"]
    }
}
