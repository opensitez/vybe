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
    two_dim_array_basic => {
        body: "int m[2][3] = {{1,2,3},{4,5,6}};\nprintf(\"%d %d\\n\", m[0][1], m[1][2]);\nreturn 0;",
        expect: ["2 6"]
    },
    two_dim_array_traversal => {
        body: r#"
int m[2][2] = {{1,2},{3,4}};
for (int i = 0; i < 2; i++) {
    for (int j = 0; j < 2; j++) {
        printf("%d\n", m[i][j]);
    }
}
return 0;
"#,
        expect: ["1", "2", "3", "4"]
    },
    two_dim_array_row_sum => {
        body: r#"
int m[3][3] = {{1,2,3},{4,5,6},{7,8,9}};
int sum = 0;
for (int j = 0; j < 3; j++) sum += m[1][j];
printf("%d\n", sum);
return 0;
"#,
        expect: ["15"]
    },
    two_dim_array_write => {
        body: r#"
int m[2][2];
m[0][0] = 10; m[0][1] = 20;
m[1][0] = 30; m[1][1] = 40;
printf("%d %d %d %d\n", m[0][0], m[0][1], m[1][0], m[1][1]);
return 0;
"#,
        expect: ["10 20 30 40"]
    },
    three_dim_array => {
        body: "int arr[2][2][2] = {{{1,2},{3,4}},{{5,6},{7,8}}};\nprintf(\"%d %d\\n\", arr[0][1][0], arr[1][0][1]);\nreturn 0;",
        expect: ["3 6"]
    },
    two_dim_char_array_strings => {
        body: "char words[3][6] = {\"one\", \"two\", \"three\"};\nprintf(\"%s %s %s\\n\", words[0], words[1], words[2]);\nreturn 0;",
        expect: ["one two three"]
    }
}

#[test]
fn two_dim_array_passed_to_function() {
    assert_program(
        &["<stdio.h>"],
        "void print_row(int row[3]) { printf(\"%d %d %d\\n\", row[0], row[1], row[2]); }",
        "int m[2][3] = {{1,2,3},{4,5,6}};\nprint_row(m[1]);\nreturn 0;",
        &["4 5 6"],
    );
}
