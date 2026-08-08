use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<string.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    strlen_of_empty_string => { declarations: "", body: "printf(\"%d\\n\", strlen(\"\"));\nreturn 0;", expect: ["0"] },
    strlen_counts_spaces => { declarations: "", body: "printf(\"%d\\n\", strlen(\"a b c\"));\nreturn 0;", expect: ["5"] },
    strlen_counts_punctuation => { declarations: "", body: "printf(\"%d\\n\", strlen(\"hi!\"));\nreturn 0;", expect: ["3"] },
    strcmp_equal_strings_returns_zero => { declarations: "", body: "printf(\"%d\\n\", strcmp(\"same\", \"same\"));\nreturn 0;", expect: ["0"] },
    strcmp_less_than_is_negative => { declarations: "", body: "if (strcmp(\"abc\", \"abd\") < 0) puts(\"lt\"); else puts(\"bad\");\nreturn 0;", expect: ["lt"] },
    strcmp_greater_than_is_positive => { declarations: "", body: "if (strcmp(\"abd\", \"abc\") > 0) puts(\"gt\"); else puts(\"bad\");\nreturn 0;", expect: ["gt"] },
    strcmp_is_case_sensitive => { declarations: "", body: "if (strcmp(\"Cat\", \"cat\") < 0) puts(\"case\"); else puts(\"bad\");\nreturn 0;", expect: ["case"] },
    strchr_returns_suffix_from_first_match => { declarations: "", body: "puts(strchr(\"banana\", 'a'));\nreturn 0;", expect: ["anana"] },
    strchr_returns_full_string_for_first_character => { declarations: "", body: "puts(strchr(\"banana\", 'b'));\nreturn 0;", expect: ["banana"] },
    strchr_returns_null_when_not_found => { declarations: "", body: "if (strchr(\"banana\", 'z') == NULL) puts(\"null\"); else puts(\"bad\");\nreturn 0;", expect: ["null"] },
    strchr_finds_space_character => { declarations: "", body: "puts(strchr(\"a b\", ' '));\nreturn 0;", expect: [" b"] },
    strchr_finds_digit_character => { declarations: "", body: "puts(strchr(\"id7\", '7'));\nreturn 0;", expect: ["7"] },
    strstr_returns_suffix_from_match => { declarations: "", body: "puts(strstr(\"bananarama\", \"nara\"));\nreturn 0;", expect: ["narama"] },
    strstr_returns_whole_string_for_prefix_match => { declarations: "", body: "puts(strstr(\"alphabet\", \"alp\"));\nreturn 0;", expect: ["alphabet"] },
    strstr_returns_null_when_missing => { declarations: "", body: "if (strstr(\"alphabet\", \"xyz\") == NULL) puts(\"null\"); else puts(\"bad\");\nreturn 0;", expect: ["null"] },
    strstr_empty_needle_returns_haystack => { declarations: "", body: "puts(strstr(\"alphabet\", \"\"));\nreturn 0;", expect: ["alphabet"] },
    strcat_appends_to_existing_prefix => { declarations: "char left[32] = \"hello\";", body: "puts(strcat(left, \" world\"));\nreturn 0;", expect: ["hello world"] },
    strcat_handles_empty_suffix => { declarations: "char left[32] = \"hello\";", body: "puts(strcat(left, \"\"));\nreturn 0;", expect: ["hello"] },
    strcat_handles_empty_prefix => { declarations: "char left[32] = \"\";", body: "puts(strcat(left, \"world\"));\nreturn 0;", expect: ["world"] },
    strcpy_overwrites_existing_buffer => { declarations: "char dest[32] = \"zzzz\";", body: "puts(strcpy(dest, \"cat\"));\nreturn 0;", expect: ["cat"] },
    strcpy_copies_empty_string => { declarations: "char dest[32] = \"zzzz\";", body: "puts(strcpy(dest, \"\"));\nreturn 0;", expect: [""] },
    strcpy_then_strlen_uses_new_contents => { declarations: "char dest[32] = \"zzzz\";", body: "strcpy(dest, \"tool\");\nprintf(\"%d\\n\", strlen(dest));\nreturn 0;", expect: ["4"] },
    strcat_result_can_be_measured => { declarations: "char left[32] = \"cat\";", body: "strcat(left, \"fish\");\nprintf(\"%d\\n\", strlen(left));\nreturn 0;", expect: ["7"] },
    strcmp_after_strcpy_matches_source => { declarations: "char dest[32] = \"aaaa\";", body: "strcpy(dest, \"copy\");\nprintf(\"%d\\n\", strcmp(dest, \"copy\"));\nreturn 0;", expect: ["0"] },
    strstr_finds_repeated_pattern_first_occurrence => { declarations: "", body: "puts(strstr(\"aaaaab\", \"aaab\"));\nreturn 0;", expect: ["aaab"] },
    strchr_finds_last_character_when_unique => { declarations: "", body: "puts(strchr(\"planet\", 't'));\nreturn 0;", expect: ["t"] },
    strcmp_empty_less_than_nonempty => { declarations: "", body: "if (strcmp(\"\", \"a\") < 0) puts(\"lt\"); else puts(\"bad\");\nreturn 0;", expect: ["lt"] },
    strcmp_nonempty_greater_than_empty => { declarations: "", body: "if (strcmp(\"a\", \"\") > 0) puts(\"gt\"); else puts(\"bad\");\nreturn 0;", expect: ["gt"] },
    strlen_after_strcat_includes_suffix => { declarations: "char left[32] = \"moon\";", body: "strcat(left, \"light\");\nprintf(\"%d\\n\", strlen(left));\nreturn 0;", expect: ["9"] },
    strcpy_preserves_embedded_spaces => { declarations: "char dest[32] = \"aaaa\";", body: "puts(strcpy(dest, \"a b\"));\nreturn 0;", expect: ["a b"] },
    strstr_match_can_start_midword => { declarations: "", body: "puts(strstr(\"mississippi\", \"issi\"));\nreturn 0;", expect: ["ississippi"] },
    strchr_can_find_null_terminator_position => { declarations: "", body: "if (strchr(\"abc\", '\\0') != NULL) puts(\"term\"); else puts(\"bad\");\nreturn 0;", expect: ["term"] },
    strcmp_numeric_text_uses_lexicographic_order => { declarations: "", body: "if (strcmp(\"10\", \"2\") < 0) puts(\"lex\"); else puts(\"bad\");\nreturn 0;", expect: ["lex"] },
    strcpy_return_value_is_destination_pointer => { declarations: "char dest[32] = \"aaaa\";", body: "if (strcpy(dest, \"tool\") == dest) puts(\"same\"); else puts(\"bad\");\nreturn 0;", expect: ["same"] },
    strcat_return_value_is_destination_pointer => { declarations: "char dest[32] = \"a\";", body: "if (strcat(dest, \"b\") == dest) puts(\"same\"); else puts(\"bad\");\nreturn 0;", expect: ["same"] },
    strstr_returns_pointer_comparable_to_haystack => { declarations: "char text[32] = \"alphabet\";", body: "if (strstr(text, \"alp\") == text) puts(\"prefix\"); else puts(\"bad\");\nreturn 0;", expect: ["prefix"] },
    strchr_result_can_be_offset => { declarations: "char text[32] = \"banana\";", body: "printf(\"%d\\n\", (int)(strchr(text, 'n') - text));\nreturn 0;", expect: ["2"] },
    strstr_result_can_be_offset => { declarations: "char text[32] = \"bananarama\";", body: "printf(\"%d\\n\", (int)(strstr(text, \"rama\") - text));\nreturn 0;", expect: ["6"] },
    strlen_of_string_literal_in_expression => { declarations: "", body: "printf(\"%d\\n\", strlen(\"vybe compiler\") + 1);\nreturn 0;", expect: ["14"] }
}
