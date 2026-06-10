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
    switch_fallthrough_multiple_cases => {
        body: r#"
int x = 2;
switch (x) {
    case 1:
    case 2:
    case 3:
        printf("1-3\n");
        break;
    default:
        printf("other\n");
}
return 0;
"#,
        expect: ["1-3"]
    },
    switch_default_in_middle => {
        body: r#"
int x = 99;
switch (x) {
    case 1: printf("one\n"); break;
    default: printf("default\n"); break;
    case 2: printf("two\n"); break;
}
return 0;
"#,
        expect: ["default"]
    },
    switch_char_value => {
        body: r#"
char c = 'b';
switch (c) {
    case 'a': printf("a\n"); break;
    case 'b': printf("b\n"); break;
    case 'c': printf("c\n"); break;
}
return 0;
"#,
        expect: ["b"]
    },
    switch_in_loop => {
        body: r#"
for (int i = 0; i < 3; i++) {
    switch (i) {
        case 0: printf("zero\n"); break;
        case 1: printf("one\n"); break;
        case 2: printf("two\n"); break;
    }
}
return 0;
"#,
        expect: ["zero", "one", "two"]
    },
    switch_fallthrough_intentional => {
        body: r#"
int x = 1;
switch (x) {
    case 1:
        printf("one ");
    case 2:
        printf("two\n");
        break;
    case 3:
        printf("three\n");
}
return 0;
"#,
        expect: ["one two"]
    },
    switch_nested_in_case => {
        body: r#"
int a = 1, b = 2;
switch (a) {
    case 1:
        switch (b) {
            case 1: printf("1,1\n"); break;
            case 2: printf("1,2\n"); break;
        }
        break;
}
return 0;
"#,
        expect: ["1,2"]
    }
}
