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
    decimal_integer_literal_prints_value => { body: "printf(\"%d\\n\", 42);\nreturn 0;", expect: ["42"] },
    hex_integer_literal_prints_decimal_value => { body: "printf(\"%d\\n\", 0x2A);\nreturn 0;", expect: ["42"] },
    octal_integer_literal_prints_decimal_value => { body: "printf(\"%d\\n\", 052);\nreturn 0;", expect: ["42"] },
    unsigned_literal_prints_with_unsigned_format => { body: "printf(\"%u\\n\", 42u);\nreturn 0;", expect: ["42"] },
    long_literal_prints_with_long_format => { body: "printf(\"%ld\\n\", 42l);\nreturn 0;", expect: ["42"] },
    float_literal_prints_fraction => { body: "printf(\"%.2f\\n\", 3.25);\nreturn 0;", expect: ["3.25"] },
    scientific_literal_prints_scaled_value => { body: "printf(\"%.1f\\n\", 1.25e2);\nreturn 0;", expect: ["125.0"] },
    uppercase_scientific_literal_prints_scaled_value => { body: "printf(\"%.1f\\n\", 1.25E2);\nreturn 0;", expect: ["125.0"] },
    char_literal_prints_character => { body: "printf(\"%c\\n\", 'A');\nreturn 0;", expect: ["A"] },
    char_literal_prints_ascii_code_when_formatted_as_int => { body: "printf(\"%d\\n\", 'A');\nreturn 0;", expect: ["65"] },
    string_literal_prints_text => { body: "puts(\"vybe\");\nreturn 0;", expect: ["vybe"] },
    adjacent_string_literals_concatenate => { body: "puts(\"vy\" \"be\");\nreturn 0;", expect: ["vybe"] },
    zero_literal_is_false_in_condition => { body: "if (0) puts(\"bad\"); else puts(\"zero\");\nreturn 0;", expect: ["zero"] },
    one_literal_is_true_in_condition => { body: "if (1) puts(\"one\"); else puts(\"bad\");\nreturn 0;", expect: ["one"] },
    negative_integer_literal_prints_sign => { body: "printf(\"%d\\n\", -42);\nreturn 0;", expect: ["-42"] },
    leading_zero_float_literal_prints_fraction => { body: "printf(\"%.2f\\n\", 0.75);\nreturn 0;", expect: ["0.75"] },
    trailing_decimal_zero_literal_is_double => { body: "printf(\"%.1f\\n\", 5.0);\nreturn 0;", expect: ["5.0"] },
    hex_literal_can_participate_in_arithmetic => { body: "printf(\"%d\\n\", 0x10 + 5);\nreturn 0;", expect: ["21"] },
    octal_literal_can_participate_in_arithmetic => { body: "printf(\"%d\\n\", 010 + 2);\nreturn 0;", expect: ["10"] },
    unsigned_hex_literal_prints_decimal_unsigned => { body: "printf(\"%u\\n\", 0xffu);\nreturn 0;", expect: ["255"] },
    long_scientific_literal_can_be_cast_to_int => { body: "printf(\"%d\\n\", (int)1.9e1);\nreturn 0;", expect: ["19"] },
    char_escape_newline_has_ascii_code_ten => { body: "printf(\"%d\\n\", '\\n');\nreturn 0;", expect: ["10"] },
    char_escape_tab_has_ascii_code_nine => { body: "printf(\"%d\\n\", '\\t');\nreturn 0;", expect: ["9"] },
    hexadecimal_escape_in_char_literal_maps_to_ascii => { body: "printf(\"%c\\n\", '\\x41');\nreturn 0;", expect: ["A"] },
    octal_escape_in_char_literal_maps_to_ascii => { body: "printf(\"%c\\n\", '\\101');\nreturn 0;", expect: ["A"] }
}