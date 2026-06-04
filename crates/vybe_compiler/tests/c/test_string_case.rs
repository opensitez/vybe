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
    strupr_converts_lowercase_word => { body: "puts(strupr(\"hello\"));\nreturn 0;", expect: ["HELLO"] },
    strupr_preserves_existing_uppercase => { body: "puts(strupr(\"Hello\"));\nreturn 0;", expect: ["HELLO"] },
    strupr_preserves_digits => { body: "puts(strupr(\"abc123\"));\nreturn 0;", expect: ["ABC123"] },
    strupr_preserves_punctuation => { body: "puts(strupr(\"hi!\"));\nreturn 0;", expect: ["HI!"] },
    strupr_handles_empty_string => { body: "puts(strupr(\"\"));\nreturn 0;", expect: [""] },
    strupr_converts_mixed_case_phrase => { body: "puts(strupr(\"MiXeD Case\"));\nreturn 0;", expect: ["MIXED CASE"] },
    strupr_allows_length_measurement => { body: "printf(\"%d\\n\", strlen(strupr(\"hello\")));\nreturn 0;", expect: ["5"] },
    strupr_result_compares_equal_to_upper_literal => { body: "if (strcmp(strupr(\"hello\"), \"HELLO\") == 0) puts(\"ok\"); else puts(\"bad\");\nreturn 0;", expect: ["ok"] },
    strlwr_converts_uppercase_word => { body: "puts(strlwr(\"HELLO\"));\nreturn 0;", expect: ["hello"] },
    strlwr_preserves_existing_lowercase => { body: "puts(strlwr(\"Hello\"));\nreturn 0;", expect: ["hello"] },
    strlwr_preserves_digits => { body: "puts(strlwr(\"ABC123\"));\nreturn 0;", expect: ["abc123"] },
    strlwr_preserves_punctuation => { body: "puts(strlwr(\"HI!\"));\nreturn 0;", expect: ["hi!"] },
    strlwr_handles_empty_string => { body: "puts(strlwr(\"\"));\nreturn 0;", expect: [""] },
    strlwr_converts_mixed_case_phrase => { body: "puts(strlwr(\"MiXeD Case\"));\nreturn 0;", expect: ["mixed case"] },
    strlwr_allows_length_measurement => { body: "printf(\"%d\\n\", strlen(strlwr(\"HELLO\")));\nreturn 0;", expect: ["5"] },
    strlwr_result_compares_equal_to_lower_literal => { body: "if (strcmp(strlwr(\"HELLO\"), \"hello\") == 0) puts(\"ok\"); else puts(\"bad\");\nreturn 0;", expect: ["ok"] },
    upper_then_lower_round_trip_ascii => { body: "puts(strlwr(strupr(\"MixEd\")));\nreturn 0;", expect: ["mixed"] },
    lower_then_upper_round_trip_ascii => { body: "puts(strupr(strlwr(\"MixEd\")));\nreturn 0;", expect: ["MIXED"] },
    strupr_result_starts_with_upper_prefix => { body: "puts(strchr(strupr(\"vybe\"), 'B'));\nreturn 0;", expect: ["BE"] },
    strlwr_result_contains_lowercase_suffix => { body: "puts(strstr(strlwr(\"HELLOWORLD\"), \"world\"));\nreturn 0;", expect: ["world"] }
}
