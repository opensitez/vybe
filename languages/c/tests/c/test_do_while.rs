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
    do_while_executes_once_when_false => {
        body: "int x = 0;\ndo { printf(\"once\\n\"); x++; } while (x < 1);\nreturn 0;",
        expect: ["once"]
    },
    do_while_executes_body_before_test => {
        body: "int x = 10;\ndo { printf(\"%d\\n\", x); x--; } while (x > 8);\nreturn 0;",
        expect: ["10", "9"]
    },
    do_while_with_break => {
        body: "int i = 0;\ndo { if (i == 2) break; printf(\"%d\\n\", i); i++; } while (i < 5);\nreturn 0;",
        expect: ["0", "1"]
    },
    do_while_counter => {
        body: "int sum = 0; int i = 1;\ndo { sum += i; i++; } while (i <= 5);\nprintf(\"%d\\n\", sum);\nreturn 0;",
        expect: ["15"]
    },
    do_while_nested => {
        body: r#"
int i = 0;
do {
    int j = 0;
    do {
        printf("%d%d\n", i, j);
        j++;
    } while (j < 2);
    i++;
} while (i < 2);
return 0;
"#,
        expect: ["00", "01", "10", "11"]
    },
    do_while_with_continue => {
        body: r#"
int i = 0;
do {
    i++;
    if (i == 2) continue;
    printf("%d\n", i);
} while (i < 4);
return 0;
"#,
        expect: ["1", "3", "4"]
    }
}
