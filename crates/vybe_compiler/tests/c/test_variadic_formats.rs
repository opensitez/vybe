use super::helpers::*;

// Edge cases in printf formatting beyond what test_printf_formats covers
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
    printf_long_long_format => {
        body: "long long x = 9876543210LL;\nprintf(\"%lld\\n\", x);\nreturn 0;",
        expect: ["9876543210"]
    },
    printf_unsigned_long_format => {
        body: "unsigned long x = 3000000000UL;\nprintf(\"%lu\\n\", x);\nreturn 0;",
        expect: ["3000000000"]
    },
    printf_size_t_format => {
        body: "size_t s = sizeof(int);\nprintf(\"%zu\\n\", s);\nreturn 0;",
        expect: ["4"]
    },
    printf_ptrdiff_format => {
        body: "int arr[3];\nptrdiff_t d = &arr[2] - &arr[0];\nprintf(\"%td\\n\", d);\nreturn 0;",
        expect: ["2"]
    },
    printf_string_null_term => {
        body: "printf(\"%s\\n\", \"hello\\0world\");\nreturn 0;",
        expect: ["hello"]
    },
    printf_hh_format_signed_char => {
        body: "signed char c = -1;\nprintf(\"%hhd\\n\", c);\nreturn 0;",
        expect: ["-1"]
    },
    printf_h_format_short => {
        body: "short s = 32767;\nprintf(\"%hd\\n\", s);\nreturn 0;",
        expect: ["32767"]
    },
    printf_star_width => {
        body: "printf(\"%*d\\n\", 6, 42);\nreturn 0;",
        expect: ["    42"]
    },
    printf_star_precision => {
        body: "printf(\"%.*f\\n\", 3, 3.14159);\nreturn 0;",
        expect: ["3.142"]
    }
}
