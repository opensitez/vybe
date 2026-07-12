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
    for_loop_step_two => {
        body: "for (int i = 0; i < 10; i += 2) printf(\"%d\\n\", i);\nreturn 0;",
        expect: ["0", "2", "4", "6", "8"]
    },
    for_loop_countdown => {
        body: "for (int i = 5; i > 0; i--) printf(\"%d\\n\", i);\nreturn 0;",
        expect: ["5", "4", "3", "2", "1"]
    },
    while_loop_digit_extraction => {
        body: r#"
int n = 1234, count = 0;
while (n > 0) { count++; n /= 10; }
printf("%d\n", count);
return 0;
"#,
        expect: ["4"]
    },
    nested_loop_multiplication_table => {
        body: r#"
for (int i = 1; i <= 2; i++) {
    for (int j = 1; j <= 3; j++) {
        printf("%d\n", i*j);
    }
}
return 0;
"#,
        expect: ["1", "2", "3", "2", "4", "6"]
    },
    continue_skips_odd => {
        body: r#"
for (int i = 0; i < 6; i++) {
    if (i % 2 != 0) continue;
    printf("%d\n", i);
}
return 0;
"#,
        expect: ["0", "2", "4"]
    },
    break_at_first_match => {
        body: r#"
int arr[] = {10, 20, 30, 40, 50};
int target = 30, idx = -1;
for (int i = 0; i < 5; i++) {
    if (arr[i] == target) { idx = i; break; }
}
printf("%d\n", idx);
return 0;
"#,
        expect: ["2"]
    },
    nested_loop_break_inner => {
        body: r#"
for (int i = 0; i < 3; i++) {
    for (int j = 0; j < 3; j++) {
        if (j == 1) break;
        printf("%d%d\n", i, j);
    }
}
return 0;
"#,
        expect: ["00", "10", "20"]
    },
    infinite_loop_with_break => {
        body: r#"
int n = 0;
while (1) {
    n++;
    if (n == 3) break;
}
printf("%d\n", n);
return 0;
"#,
        expect: ["3"]
    },
    for_loop_multiple_init_update => {
        body: "int i, j;\nfor (i=0, j=10; i<5; i++, j--) printf(\"%d %d\\n\", i, j);\nreturn 0;",
        expect: ["0 10", "1 9", "2 8", "3 7", "4 6"]
    }
}
