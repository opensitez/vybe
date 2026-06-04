use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<string.h>", "<stdlib.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    strcpy_can_copy_shorter_text_over_longer_buffer => { declarations: "char dest[32] = \"alphabet\";", body: "puts(strcpy(dest, \"cat\")); return 0;", expect: ["cat"] },
    strcpy_can_copy_same_length_text => { declarations: "char dest[32] = \"abc\";", body: "puts(strcpy(dest, \"xyz\")); return 0;", expect: ["xyz"] },
    strcpy_can_copy_text_with_spaces => { declarations: "char dest[32] = \"abc\";", body: "puts(strcpy(dest, \"a b c\")); return 0;", expect: ["a b c"] },
    strcpy_can_copy_text_with_punctuation => { declarations: "char dest[32] = \"abc\";", body: "puts(strcpy(dest, \"hi!\")); return 0;", expect: ["hi!"] },
    strcpy_then_strcat_can_append_additional_text => { declarations: "char dest[32] = \"abc\";", body: "strcpy(dest, \"hi\"); strcat(dest, \" there\"); puts(dest); return 0;", expect: ["hi there"] },
    strcat_can_append_single_character_string => { declarations: "char dest[32] = \"a\";", body: "puts(strcat(dest, \"b\")); return 0;", expect: ["ab"] },
    strcat_can_append_phrase_with_spaces => { declarations: "char dest[32] = \"hello\";", body: "puts(strcat(dest, \" world\")); return 0;", expect: ["hello world"] },
    strcat_can_append_punctuation => { declarations: "char dest[32] = \"hello\";", body: "puts(strcat(dest, \"!\")); return 0;", expect: ["hello!"] },
    strlen_after_strcpy_reflects_new_length => { declarations: "char dest[32] = \"alphabet\";", body: "strcpy(dest, \"cat\"); printf(\"%d\\n\", strlen(dest)); return 0;", expect: ["3"] },
    strlen_after_strcat_reflects_combined_length => { declarations: "char dest[32] = \"cat\";", body: "strcat(dest, \"fish\"); printf(\"%d\\n\", strlen(dest)); return 0;", expect: ["7"] },
    strcmp_after_copy_can_confirm_equality => { declarations: "char dest[32] = \"abc\";", body: "strcpy(dest, \"copy\"); printf(\"%d\\n\", strcmp(dest, \"copy\")); return 0;", expect: ["0"] },
    strcmp_after_concat_can_compare_lexicographically => { declarations: "char dest[32] = \"ab\";", body: "strcat(dest, \"c\"); if (strcmp(dest, \"abd\") < 0) puts(\"lt\"); else puts(\"bad\"); return 0;", expect: ["lt"] },
    strchr_after_copy_can_find_first_character => { declarations: "char dest[32] = \"abc\";", body: "strcpy(dest, \"tool\"); puts(strchr(dest, 't')); return 0;", expect: ["tool"] },
    strstr_after_concat_can_find_suffix => { declarations: "char dest[32] = \"ab\";", body: "strcat(dest, \"cdef\"); puts(strstr(dest, \"def\")); return 0;", expect: ["def"] },
    strcpy_return_value_can_be_chained_into_strcat => { declarations: "char dest[32] = \"\";", body: "strcat(strcpy(dest, \"cat\"), \"fish\"); puts(dest); return 0;", expect: ["catfish"] },
    strcat_can_be_chained_multiple_times => { declarations: "char dest[32] = \"a\";", body: "strcat(dest, \"b\"); strcat(dest, \"c\"); puts(dest); return 0;", expect: ["abc"] },
    strcpy_can_copy_empty_string_into_buffer => { declarations: "char dest[32] = \"seed\";", body: "puts(strcpy(dest, \"\")); return 0;", expect: [""] },
    strcat_with_empty_suffix_leaves_text_unchanged => { declarations: "char dest[32] = \"seed\";", body: "puts(strcat(dest, \"\")); return 0;", expect: ["seed"] },
    strcat_after_empty_copy_produces_suffix_only => { declarations: "char dest[32] = \"seed\";", body: "strcpy(dest, \"\"); strcat(dest, \"tail\"); puts(dest); return 0;", expect: ["tail"] },
    strcpy_can_copy_numeric_text_for_atoi => { declarations: "char dest[32] = \"\";", body: "strcpy(dest, \"42\"); printf(\"%d\\n\", atoi(dest)); return 0;", expect: ["42"] },
    strcat_can_build_path_like_string => { declarations: "char dest[32] = \"usr\";", body: "strcat(dest, \"/bin\"); puts(dest); return 0;", expect: ["usr/bin"] },
    strcpy_can_copy_escape_sequences_as_text => { declarations: "char dest[32] = \"\";", body: "puts(strcpy(dest, \"a\\tb\")); return 0;", expect: ["a\tb"] },
    strcmp_can_detect_prefix_after_copy => { declarations: "char dest[32] = \"\";", body: "strcpy(dest, \"alpha\"); if (strcmp(dest, \"alphabet\") < 0) puts(\"lt\"); else puts(\"bad\"); return 0;", expect: ["lt"] },
    strchr_can_find_punctuation_after_concat => { declarations: "char dest[32] = \"hi\";", body: "strcat(dest, \"!\"); puts(strchr(dest, '!')); return 0;", expect: ["!"] },
    strstr_can_find_middle_segment_after_copy => { declarations: "char dest[32] = \"\";", body: "strcpy(dest, \"bananarama\"); puts(strstr(dest, \"nara\")); return 0;", expect: ["narama"] },
    strcpy_overwrites_old_suffix_completely => { declarations: "char dest[32] = \"abcdef\";", body: "strcpy(dest, \"xy\"); puts(dest); return 0;", expect: ["xy"] },
    strcat_preserves_existing_prefix_when_appending => { declarations: "char dest[32] = \"pre\";", body: "strcat(dest, \"fix\"); puts(dest); return 0;", expect: ["prefix"] },
    strcmp_between_copied_and_original_buffer_can_match => { declarations: "char left[32] = \"copy\"; char right[32] = \"\";", body: "strcpy(right, left); printf(\"%d\\n\", strcmp(left, right)); return 0;", expect: ["0"] },
    strlen_can_measure_buffer_after_multiple_concats => { declarations: "char dest[32] = \"a\";", body: "strcat(dest, \"b\"); strcat(dest, \"cd\"); printf(\"%d\\n\", strlen(dest)); return 0;", expect: ["4"] },
    strcpy_and_puts_can_round_trip_single_character => { declarations: "char dest[32] = \"\";", body: "puts(strcpy(dest, \"z\")); return 0;", expect: ["z"] },
    strcat_can_append_digit_text => { declarations: "char dest[32] = \"id\";", body: "puts(strcat(dest, \"7\")); return 0;", expect: ["id7"] },
    strstr_after_copy_can_report_not_found_via_null => { declarations: "char dest[32] = \"\";", body: "strcpy(dest, \"alpha\"); if (strstr(dest, \"zzz\") == NULL) puts(\"null\"); else puts(\"bad\"); return 0;", expect: ["null"] },
    strchr_after_copy_can_report_not_found_via_null => { declarations: "char dest[32] = \"\";", body: "strcpy(dest, \"alpha\"); if (strchr(dest, 'z') == NULL) puts(\"null\"); else puts(\"bad\"); return 0;", expect: ["null"] },
    strcat_return_value_can_feed_strlen => { declarations: "char dest[32] = \"a\";", body: "printf(\"%d\\n\", strlen(strcat(dest, \"bc\"))); return 0;", expect: ["3"] },
    strcpy_return_value_can_feed_strlen => { declarations: "char dest[32] = \"seed\";", body: "printf(\"%d\\n\", strlen(strcpy(dest, \"cat\"))); return 0;", expect: ["3"] },
    copy_then_concat_then_search_can_find_tail => { declarations: "char dest[32] = \"\";", body: "strcpy(dest, \"hello\"); strcat(dest, \"world\"); puts(strstr(dest, \"world\")); return 0;", expect: ["world"] },
    copy_then_concat_then_compare_can_equal_expected => { declarations: "char dest[32] = \"\";", body: "strcpy(dest, \"hello\"); strcat(dest, \"world\"); printf(\"%d\\n\", strcmp(dest, \"helloworld\")); return 0;", expect: ["0"] },
    strcat_with_space_separator_builds_phrase => { declarations: "char dest[32] = \"hello\";", body: "strcat(dest, \" \" ); strcat(dest, \"vybe\"); puts(dest); return 0;", expect: ["hello vybe"] },
    strcpy_then_character_index_reads_expected_value => { declarations: "char dest[32] = \"\";", body: "strcpy(dest, \"cat\"); printf(\"%c\\n\", dest[1]); return 0;", expect: ["a"] },
    strcat_can_append_to_explicit_null_terminated_buffer => { declarations: "char dest[32] = {'h', 'i', '\\0'};", body: "puts(strcat(dest, \"!\")); return 0;", expect: ["hi!"] },
    strcpy_can_copy_mixed_case_text => { declarations: "char dest[32] = \"\";", body: "puts(strcpy(dest, \"MiXeD\")); return 0;", expect: ["MiXeD"] }
}
