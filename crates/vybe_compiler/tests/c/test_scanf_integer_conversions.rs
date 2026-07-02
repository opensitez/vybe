//! sscanf integer conversions — %d %i %u %o %x %X with distinct inputs.


c_run_cases! {
    sscanf_d_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"0\", \"%d\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["0"]
    },
    sscanf_d_single_digit_seven => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"7\", \"%d\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["7"]
    },
    sscanf_d_leading_whitespace => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"  15\", \"%d\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["15"]
    },
    sscanf_d_plus_sign_prefix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"+33\", \"%d\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["33"]
    },
    sscanf_d_negative_ninety_nine => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"-99\", \"%d\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["-99"]
    },
    sscanf_d_large_positive => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"12345\", \"%d\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["12345"]
    },
    sscanf_d_stops_at_nondigit => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"88xy\", \"%d\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["88"]
    },
    sscanf_d_field_width_three => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"12345\", \"%3d\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["123"]
    },
    sscanf_d_two_values_sequential => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a,b; sscanf(\"3 4\", \"%d %d\", &a, &b); printf(\"%d %d\\n\", a, b); return 0;",
        expect: ["3 4"]
    },
    sscanf_d_assignment_count => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; int c=sscanf(\"9\", \"%d\", &n); printf(\"%d %d\\n\", c, n); return 0;",
        expect: ["1 9"]
    },
    sscanf_d_tab_prefixed => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"\\t21\", \"%d\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["21"]
    },
    sscanf_d_mixed_sign_pair => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a,b; sscanf(\"-1 +2\", \"%d %d\", &a, &b); printf(\"%d %d\\n\", a, b); return 0;",
        expect: ["-1 2"]
    },
    sscanf_d_literal_comma_separator => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a,b; sscanf(\"6,7\", \"%d,%d\", &a, &b); printf(\"%d %d\\n\", a, b); return 0;",
        expect: ["6 7"]
    },
    sscanf_d_width_one_digit => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"987\", \"%1d\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["9"]
    },
    sscanf_d_three_digit_width => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"54321\", \"%3d\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["543"]
    },
    sscanf_i_plain_decimal_nineteen => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"19\", \"%i\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["19"]
    },
    sscanf_i_leading_zeros_decimal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"007\", \"%i\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["7"]
    },
    sscanf_i_hex_lower_prefix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"0x1f\", \"%i\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["31"]
    },
    sscanf_i_hex_upper_prefix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"0XAB\", \"%i\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["171"]
    },
    sscanf_i_octal_leading_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"077\", \"%i\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["63"]
    },
    sscanf_i_negative_fifty => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"-50\", \"%i\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["-50"]
    },
    sscanf_i_negative_hex => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"-0x10\", \"%i\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["-16"]
    },
    sscanf_i_plus_hundred => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"+100\", \"%i\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["100"]
    },
    sscanf_i_zero_alone => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"0\", \"%i\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["0"]
    },
    sscanf_i_width_four_digits => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"123456\", \"%4i\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["1234"]
    },
    sscanf_i_hex_then_decimal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a,b; sscanf(\"0x10 8\", \"%i %i\", &a, &b); printf(\"%d %d\\n\", a, b); return 0;",
        expect: ["16 8"]
    },
    sscanf_i_octal_then_decimal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a,b; sscanf(\"010 10\", \"%i %i\", &a, &b); printf(\"%d %d\\n\", a, b); return 0;",
        expect: ["8 10"]
    },
    sscanf_i_decimal_after_whitespace_run => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; sscanf(\"   256\", \"%i\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["256"]
    },
    sscanf_u_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned u; sscanf(\"0\", \"%u\", &u); printf(\"%u\\n\", u); return 0;",
        expect: ["0"]
    },
    sscanf_u_fifty => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned u; sscanf(\"50\", \"%u\", &u); printf(\"%u\\n\", u); return 0;",
        expect: ["50"]
    },
    sscanf_u_large_32bit_max => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned u; sscanf(\"4294967295\", \"%u\", &u); printf(\"%u\\n\", u); return 0;",
        expect: ["4294967295"]
    },
    sscanf_u_leading_spaces => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned u; sscanf(\"  200\", \"%u\", &u); printf(\"%u\\n\", u); return 0;",
        expect: ["200"]
    },
    sscanf_u_width_two_digits => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned u; sscanf(\"789\", \"%2u\", &u); printf(\"%u\\n\", u); return 0;",
        expect: ["78"]
    },
    sscanf_u_pair_values => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned a,b; sscanf(\"10 20\", \"%u %u\", &a, &b); printf(\"%u %u\\n\", a, b); return 0;",
        expect: ["10 20"]
    },
    sscanf_u_trailing_alpha => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned u; sscanf(\"300rest\", \"%u\", &u); printf(\"%u\\n\", u); return 0;",
        expect: ["300"]
    },
    sscanf_u_million => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned u; sscanf(\"1000000\", \"%u\", &u); printf(\"%u\\n\", u); return 0;",
        expect: ["1000000"]
    },
    sscanf_u_assignment_count => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned u; int c=sscanf(\"17\", \"%u\", &u); printf(\"%d %u\\n\", c, u); return 0;",
        expect: ["1 17"]
    },
    sscanf_u_three_value_run => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned a,b,c; sscanf(\"1 2 3\", \"%u %u %u\", &a, &b, &c); printf(\"%u %u %u\\n\", a, b, c); return 0;",
        expect: ["1 2 3"]
    },
    sscanf_o_seven_octal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned o; sscanf(\"7\", \"%o\", &o); printf(\"%u\\n\", o); return 0;",
        expect: ["7"]
    },
    sscanf_o_twelve_octal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned o; sscanf(\"14\", \"%o\", &o); printf(\"%u\\n\", o); return 0;",
        expect: ["12"]
    },
    sscanf_o_sixty_three_octal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned o; sscanf(\"77\", \"%o\", &o); printf(\"%u\\n\", o); return 0;",
        expect: ["63"]
    },
    sscanf_o_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned o; sscanf(\"0\", \"%o\", &o); printf(\"%u\\n\", o); return 0;",
        expect: ["0"]
    },
    sscanf_o_width_two => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned o; sscanf(\"177\", \"%2o\", &o); printf(\"%u\\n\", o); return 0;",
        expect: ["15"]
    },
    sscanf_o_pair => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned a,b; sscanf(\"10 20\", \"%o %o\", &a, &b); printf(\"%u %u\\n\", a, b); return 0;",
        expect: ["8 16"]
    },
    sscanf_x_lowercase_ff => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned x; sscanf(\"ff\", \"%x\", &x); printf(\"%u\\n\", x); return 0;",
        expect: ["255"]
    },
    sscanf_x_lowercase_dead => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned x; sscanf(\"dead\", \"%x\", &x); printf(\"%u\\n\", x); return 0;",
        expect: ["57005"]
    },
    sscanf_x_uppercase_ab => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned x; sscanf(\"AB\", \"%X\", &x); printf(\"%u\\n\", x); return 0;",
        expect: ["171"]
    },
    sscanf_x_uppercase_beef => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned x; sscanf(\"BEEF\", \"%X\", &x); printf(\"%u\\n\", x); return 0;",
        expect: ["48879"]
    },
    sscanf_x_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned x; sscanf(\"0\", \"%x\", &x); printf(\"%u\\n\", x); return 0;",
        expect: ["0"]
    },
    sscanf_x_width_three => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned x; sscanf(\"abcd\", \"%3x\", &x); printf(\"%u\\n\", x); return 0;",
        expect: ["2748"]
    },
    sscanf_x_mixed_case_pair => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned a,b; sscanf(\"a B\", \"%x %X\", &a, &b); printf(\"%u %u\\n\", a, b); return 0;",
        expect: ["10 11"]
    },
}
