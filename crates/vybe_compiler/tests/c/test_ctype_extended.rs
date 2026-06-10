use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<ctype.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    isblank_space_and_tab => {
        body: "printf(\"%d %d %d\\n\", isblank(' ') != 0, isblank('\\t') != 0, isblank('a') != 0);\nreturn 0;",
        expect: ["1 1 0"]
    },
    isgraph_excludes_space => {
        body: "printf(\"%d %d\\n\", isgraph('a') != 0, isgraph(' ') != 0);\nreturn 0;",
        expect: ["1 0"]
    },
    isprint_includes_space => {
        body: "printf(\"%d %d\\n\", isprint(' ') != 0, isprint('a') != 0);\nreturn 0;",
        expect: ["1 1"]
    },
    ispunct_marks => {
        body: "printf(\"%d %d %d\\n\", ispunct('.') != 0, ispunct(',') != 0, ispunct('a') != 0);\nreturn 0;",
        expect: ["1 1 0"]
    },
    iscntrl_control_chars => {
        body: "printf(\"%d %d\\n\", iscntrl('\\n') != 0, iscntrl('a') != 0);\nreturn 0;",
        expect: ["1 0"]
    },
    isxdigit_hex_chars => {
        body: "printf(\"%d %d %d %d\\n\", isxdigit('0') != 0, isxdigit('a') != 0, isxdigit('F') != 0, isxdigit('g') != 0);\nreturn 0;",
        expect: ["1 1 1 0"]
    },
    toupper_lowercase => {
        body: "printf(\"%c %c\\n\", toupper('a'), toupper('z'));\nreturn 0;",
        expect: ["A Z"]
    },
    tolower_uppercase => {
        body: "printf(\"%c %c\\n\", tolower('A'), tolower('Z'));\nreturn 0;",
        expect: ["a z"]
    },
    toupper_nonalpha_unchanged => {
        body: "printf(\"%c\\n\", toupper('5'));\nreturn 0;",
        expect: ["5"]
    },
    isalnum_combined => {
        body: "printf(\"%d %d %d\\n\", isalnum('a') != 0, isalnum('9') != 0, isalnum('@') != 0);\nreturn 0;",
        expect: ["1 1 0"]
    }
}
