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

#[test]
fn pointer_to_first_element() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int arr[3] = {10,20,30};\nint *p = &arr[0];\nprintf(\"%d %d %d\\n\", p[0], p[1], p[2]);\nreturn 0;",
        &["10 20 30"],
    );
}

#[test]
fn pointer_decay_from_array_name() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int arr[3] = {5,6,7};\nint *p = arr;\nprintf(\"%d %d\\n\", *p, *(p+2));\nreturn 0;",
        &["5 7"],
    );
}

#[test]
fn pointer_walk_through_array() {
    assert_program(
        &["<stdio.h>"],
        "",
        r#"
int arr[4] = {1,2,3,4};
int *p = arr;
int *end = arr + 4;
int sum = 0;
while (p < end) sum += *p++;
printf("%d\n", sum);
return 0;
"#,
        &["10"],
    );
}

c_cases! {
    pointer_to_array_2d_row => {
        declarations: "",
        body: r#"
int m[3][4] = {{1,2,3,4},{5,6,7,8},{9,10,11,12}};
int (*row)[4] = &m[1];
printf("%d %d\n", (*row)[0], (*row)[3]);
return 0;
"#,
        expect: ["5 8"]
    },
    string_pointer_array_traverse => {
        declarations: "",
        body: r#"
const char *words[] = {"one","two","three",NULL};
const char **p = words;
while (*p) { printf("%s\n", *p); p++; }
return 0;
"#,
        expect: ["one", "two", "three"]
    }
}
