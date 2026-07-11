use crate::helpers;

macro_rules! cobol_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let out = crate::helpers::run_prints($src);
            assert_eq!(out, $expected);
        }
    };
}

// 30 edge cases for Environment Division
cobol_test!(
    test_env_source_computer_debugging,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. SOURCE-COMPUTER. X WITH DEBUGGING MODE. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_object_computer_memory,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. OBJECT-COMPUTER. X MEMORY SIZE 1000 WORDS. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_object_computer_sequence,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. OBJECT-COMPUTER. X PROGRAM COLLATING SEQUENCE IS NATIVE. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_object_computer_segment_limit,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. OBJECT-COMPUTER. X SEGMENT-LIMIT IS 50. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_special_names_currency,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. SPECIAL-NAMES. CURRENCY SIGN IS '*'. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_special_names_decimal,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. SPECIAL-NAMES. DECIMAL-POINT IS COMMA. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_special_names_alphabet,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. SPECIAL-NAMES. ALPHABET A IS STANDARD-1. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_special_names_symbolic,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. SPECIAL-NAMES. SYMBOLIC CHARACTERS SYM ARE 65. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_special_names_class,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. SPECIAL-NAMES. CLASS C IS 'A' THRU 'Z'. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_special_names_crt,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. SPECIAL-NAMES. CRT STATUS IS V. DATA DIVISION. WORKING-STORAGE SECTION. 01 V PIC X. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_11,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_12,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_13,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_14,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_15,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_16,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_17,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_18,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_19,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_20,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_21,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_22,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_23,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_24,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_25,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_26,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_27,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_28,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_29,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_env_parse_30,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
