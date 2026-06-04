use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<string.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    newline_escape_creates_line_break_in_string => { body: "puts(\"a\\nb\");\nreturn 0;", expect: ["a", "b"] },
    tab_escape_is_preserved_in_output => { body: "puts(\"a\\tb\");\nreturn 0;", expect: ["a\tb"] },
    backslash_escape_preserves_literal_backslash => { body: "puts(\"a\\\\b\");\nreturn 0;", expect: ["a\\b"] },
    double_quote_escape_preserves_quote => { body: "puts(\"say \\\"hi\\\"\");\nreturn 0;", expect: ["say \"hi\""] },
    single_quote_escape_in_char_literal_maps_to_ascii => { body: "printf(\"%d\\n\", '\\\'');\nreturn 0;", expect: ["39"] },
    carriage_return_escape_has_ascii_code_thirteen => { body: "printf(\"%d\\n\", '\\r');\nreturn 0;", expect: ["13"] },
    form_feed_escape_has_ascii_code_twelve => { body: "printf(\"%d\\n\", '\\f');\nreturn 0;", expect: ["12"] },
    vertical_tab_escape_has_ascii_code_eleven => { body: "printf(\"%d\\n\", '\\v');\nreturn 0;", expect: ["11"] },
    bell_escape_has_ascii_code_seven => { body: "printf(\"%d\\n\", '\\a');\nreturn 0;", expect: ["7"] },
    question_mark_escape_preserves_question_mark => { body: "puts(\"what\\?\");\nreturn 0;", expect: ["what?"] },
    hex_escape_in_string_produces_character => { body: "puts(\"\\x41\");\nreturn 0;", expect: ["A"] },
    octal_escape_in_string_produces_character => { body: "puts(\"\\101\");\nreturn 0;", expect: ["A"] },
    escaped_null_character_terminates_visible_length => { body: "printf(\"%d\\n\", strlen(\"a\\0b\"));\nreturn 0;", expect: ["1"] },
    escaped_backslash_and_quote_can_coexist => { body: "puts(\"\\\\\"\");\nreturn 0;", expect: ["\\\""] },
    multiple_escaped_characters_keep_order => { body: "puts(\"A\\nB\\tC\");\nreturn 0;", expect: ["A", "B\tC"] },
    char_array_with_escape_sequence_prints_correctly => { body: "char text[] = \"line\\nnext\"; puts(text);\nreturn 0;", expect: ["line", "next"] },
    escaped_tab_counts_as_single_character_in_strlen => { body: "printf(\"%d\\n\", strlen(\"a\\tb\"));\nreturn 0;", expect: ["3"] },
    escaped_backslash_counts_as_single_character_in_strlen => { body: "printf(\"%d\\n\", strlen(\"a\\\\b\"));\nreturn 0;", expect: ["3"] },
    escaped_quote_counts_as_single_character_in_strlen => { body: "printf(\"%d\\n\", strlen(\"\\\"\"));\nreturn 0;", expect: ["1"] },
    escape_sequence_can_be_used_in_character_comparison => { body: "if ('\\n' == 10) puts(\"match\"); else puts(\"bad\");\nreturn 0;", expect: ["match"] },
    hex_escape_can_feed_char_format => { body: "printf(\"%c\\n\", '\\x42');\nreturn 0;", expect: ["B"] },
    octal_escape_can_feed_char_format => { body: "printf(\"%c\\n\", '\\102');\nreturn 0;", expect: ["B"] },
    escaped_newline_in_char_array_keeps_null_terminated_text => { body: "char text[] = \"hi\\n\"; printf(\"%d\\n\", strlen(text));\nreturn 0;", expect: ["3"] },
    escaped_zero_char_literal_has_zero_value => { body: "printf(\"%d\\n\", '\\0');\nreturn 0;", expect: ["0"] },
    mixed_escape_literal_prints_expected_ascii_codes => { body: "printf(\"%d %d %d\\n\", '\\n', '\\t', '\\\\');\nreturn 0;", expect: ["10 9 92"] }
}
