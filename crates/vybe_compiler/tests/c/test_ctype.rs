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
    isalpha_accepts_lowercase_letter => { body: "printf(\"%d\\n\", isalpha('a') != 0); return 0;", expect: ["1"] },
    isalpha_accepts_uppercase_letter => { body: "printf(\"%d\\n\", isalpha('Z') != 0); return 0;", expect: ["1"] },
    isalpha_rejects_digit => { body: "printf(\"%d\\n\", isalpha('7') != 0); return 0;", expect: ["0"] },
    isdigit_accepts_digit => { body: "printf(\"%d\\n\", isdigit('7') != 0); return 0;", expect: ["1"] },
    isdigit_rejects_letter => { body: "printf(\"%d\\n\", isdigit('a') != 0); return 0;", expect: ["0"] },
    isalnum_accepts_letter => { body: "printf(\"%d\\n\", isalnum('a') != 0); return 0;", expect: ["1"] },
    isalnum_accepts_digit => { body: "printf(\"%d\\n\", isalnum('7') != 0); return 0;", expect: ["1"] },
    isalnum_rejects_punctuation => { body: "printf(\"%d\\n\", isalnum('!') != 0); return 0;", expect: ["0"] },
    isspace_accepts_space => { body: "printf(\"%d\\n\", isspace(' ') != 0); return 0;", expect: ["1"] },
    isspace_accepts_newline => { body: "printf(\"%d\\n\", isspace('\\n') != 0); return 0;", expect: ["1"] },
    isspace_rejects_letter => { body: "printf(\"%d\\n\", isspace('a') != 0); return 0;", expect: ["0"] },
    isupper_accepts_uppercase_letter => { body: "printf(\"%d\\n\", isupper('Q') != 0); return 0;", expect: ["1"] },
    isupper_rejects_lowercase_letter => { body: "printf(\"%d\\n\", isupper('q') != 0); return 0;", expect: ["0"] },
    islower_accepts_lowercase_letter => { body: "printf(\"%d\\n\", islower('q') != 0); return 0;", expect: ["1"] },
    islower_rejects_uppercase_letter => { body: "printf(\"%d\\n\", islower('Q') != 0); return 0;", expect: ["0"] },
    isxdigit_accepts_decimal_digit => { body: "printf(\"%d\\n\", isxdigit('9') != 0); return 0;", expect: ["1"] },
    isxdigit_accepts_lowercase_hex_letter => { body: "printf(\"%d\\n\", isxdigit('a') != 0); return 0;", expect: ["1"] },
    isxdigit_accepts_uppercase_hex_letter => { body: "printf(\"%d\\n\", isxdigit('F') != 0); return 0;", expect: ["1"] },
    isxdigit_rejects_non_hex_letter => { body: "printf(\"%d\\n\", isxdigit('g') != 0); return 0;", expect: ["0"] },
    ispunct_accepts_punctuation => { body: "printf(\"%d\\n\", ispunct('!') != 0); return 0;", expect: ["1"] },
    ispunct_rejects_letter => { body: "printf(\"%d\\n\", ispunct('a') != 0); return 0;", expect: ["0"] },
    isprint_accepts_space => { body: "printf(\"%d\\n\", isprint(' ') != 0); return 0;", expect: ["1"] },
    isprint_accepts_letter => { body: "printf(\"%d\\n\", isprint('a') != 0); return 0;", expect: ["1"] },
    iscntrl_accepts_newline => { body: "printf(\"%d\\n\", iscntrl('\\n') != 0); return 0;", expect: ["1"] },
    iscntrl_rejects_printable_letter => { body: "printf(\"%d\\n\", iscntrl('a') != 0); return 0;", expect: ["0"] },
    toupper_converts_lowercase_letter => { body: "printf(\"%c\\n\", toupper('a')); return 0;", expect: ["A"] },
    toupper_leaves_uppercase_letter_unchanged => { body: "printf(\"%c\\n\", toupper('A')); return 0;", expect: ["A"] },
    tolower_converts_uppercase_letter => { body: "printf(\"%c\\n\", tolower('A')); return 0;", expect: ["a"] },
    tolower_leaves_lowercase_letter_unchanged => { body: "printf(\"%c\\n\", tolower('a')); return 0;", expect: ["a"] },
    classification_results_can_drive_branch => { body: "if (isdigit('8') && !isalpha('8')) puts(\"digit\"); else puts(\"bad\"); return 0;", expect: ["digit"] }
}