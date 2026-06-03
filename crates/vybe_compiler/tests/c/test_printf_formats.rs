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
    printf_formats_signed_decimal => { body: "printf(\"%d\\n\", 42);\nreturn 0;", expect: ["42"] },
    printf_formats_negative_decimal => { body: "printf(\"%d\\n\", -42);\nreturn 0;", expect: ["-42"] },
    printf_formats_unsigned_decimal => { body: "printf(\"%u\\n\", 42u);\nreturn 0;", expect: ["42"] },
    printf_formats_lower_hex => { body: "printf(\"%x\\n\", 255);\nreturn 0;", expect: ["ff"] },
    printf_formats_upper_hex => { body: "printf(\"%X\\n\", 255);\nreturn 0;", expect: ["FF"] },
    printf_formats_octal => { body: "printf(\"%o\\n\", 64);\nreturn 0;", expect: ["100"] },
    printf_formats_character => { body: "printf(\"%c\\n\", 65);\nreturn 0;", expect: ["A"] },
    printf_formats_string => { body: "printf(\"%s\\n\", \"vybe\");\nreturn 0;", expect: ["vybe"] },
    printf_formats_percent_literal => { body: "printf(\"100%%\\n\");\nreturn 0;", expect: ["100%"] },
    printf_formats_float_default_precision => { body: "printf(\"%f\\n\", 3.5);\nreturn 0;", expect: ["3.500000"] },
    printf_formats_float_precision_two => { body: "printf(\"%.2f\\n\", 3.14159);\nreturn 0;", expect: ["3.14"] },
    printf_formats_float_width_and_precision => { body: "printf(\"%8.2f\\n\", 3.5);\nreturn 0;", expect: ["    3.50"] },
    printf_formats_scientific_lower => { body: "printf(\"%.2e\\n\", 1250.0);\nreturn 0;", expect: ["1.25e+3"] },
    printf_formats_scientific_upper => { body: "printf(\"%.2E\\n\", 1250.0);\nreturn 0;", expect: ["1.25E+3"] },
    printf_formats_general_lower => { body: "printf(\"%.4g\\n\", 12.3456);\nreturn 0;", expect: ["12.35"] },
    printf_formats_general_upper => { body: "printf(\"%.3G\\n\", 1250.0);\nreturn 0;", expect: ["1.25E+3"] },
    printf_formats_zero_padded_integer => { body: "printf(\"%05d\\n\", 42);\nreturn 0;", expect: ["00042"] },
    printf_formats_left_justified_integer => { body: "printf(\"%-5dX\\n\", 42);\nreturn 0;", expect: ["42   X"] },
    printf_formats_left_justified_string => { body: "printf(\"%-6s!\\n\", \"hi\");\nreturn 0;", expect: ["hi    !"] },
    printf_formats_positive_sign_flag => { body: "printf(\"%+d\\n\", 42);\nreturn 0;", expect: ["+42"] },
    printf_formats_space_sign_flag => { body: "printf(\"% d\\n\", 42);\nreturn 0;", expect: [" 42"] },
    printf_formats_alternate_hex => { body: "printf(\"%#x\\n\", 255);\nreturn 0;", expect: ["0xff"] },
    printf_formats_alternate_upper_hex => { body: "printf(\"%#X\\n\", 255);\nreturn 0;", expect: ["0XFF"] },
    printf_formats_alternate_octal => { body: "printf(\"%#o\\n\", 64);\nreturn 0;", expect: ["0100"] },
    printf_formats_multiple_placeholders => { body: "printf(\"%s:%d:%c\\n\", \"item\", 7, 65);\nreturn 0;", expect: ["item:7:A"] },
    printf_formats_width_with_string => { body: "printf(\"%6s\\n\", \"cat\");\nreturn 0;", expect: ["   cat"] },
    printf_formats_precision_truncates_string => { body: "printf(\"%.3s\\n\", \"alphabet\");\nreturn 0;", expect: ["alp"] },
    printf_formats_width_and_string_precision => { body: "printf(\"%7.3s\\n\", \"alphabet\");\nreturn 0;", expect: ["    alp"] },
    printf_formats_negative_float_sign => { body: "printf(\"%+.1f\\n\", -3.5);\nreturn 0;", expect: ["-3.5"] },
    printf_formats_positive_float_sign => { body: "printf(\"%+.1f\\n\", 3.5);\nreturn 0;", expect: ["+3.5"] },
    printf_formats_zero_padded_with_sign => { body: "printf(\"%+06d\\n\", 42);\nreturn 0;", expect: ["+00042"] },
    printf_formats_negative_width_number => { body: "printf(\"%6d\\n\", -42);\nreturn 0;", expect: ["   -42"] },
    printf_formats_long_decimal => { body: "printf(\"%ld\\n\", 123456l);\nreturn 0;", expect: ["123456"] },
    printf_formats_long_hex => { body: "printf(\"%lx\\n\", 4095l);\nreturn 0;", expect: ["fff"] },
    printf_formats_char_from_expression => { body: "printf(\"%c\\n\", 'A' + 2);\nreturn 0;", expect: ["C"] },
    printf_formats_string_and_percent_mix => { body: "printf(\"%s %% done\\n\", \"50\");\nreturn 0;", expect: ["50 % done"] },
    printf_formats_integer_zero_precision_zero => { body: "printf(\"%.0d\\n\", 0);\nreturn 0;", expect: [""] },
    printf_formats_float_rounding_half_up_display => { body: "printf(\"%.2f\\n\", 2.675);\nreturn 0;", expect: ["2.68"] },
    printf_formats_width_precision_and_sign => { body: "printf(\"%+8.2f\\n\", 12.5);\nreturn 0;", expect: ["  +12.50"] },
    printf_formats_octal_zero_value => { body: "printf(\"%#o\\n\", 0);\nreturn 0;", expect: ["0"] },
    printf_formats_hex_zero_value => { body: "printf(\"%#x\\n\", 0);\nreturn 0;", expect: ["0"] }
}