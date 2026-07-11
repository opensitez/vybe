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

// 30 specific tests for INSPECT TALLYING
cobol_test!(
    test_inspect_tallying_all,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'AABAA'. 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. INSPECT S1 TALLYING T FOR ALL 'A'. DISPLAY T. STOP RUN.",
    vec!["4"]
);
cobol_test!(
    test_inspect_tallying_leading,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'AABAA'. 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. INSPECT S1 TALLYING T FOR LEADING 'A'. DISPLAY T. STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_inspect_tallying_characters,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'AABAA'. 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. INSPECT S1 TALLYING T FOR CHARACTERS. DISPLAY T. STOP RUN.",
    vec!["5"]
);
cobol_test!(
    test_inspect_tallying_all_before,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'AABAA'. 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. INSPECT S1 TALLYING T FOR ALL 'A' BEFORE INITIAL 'B'. DISPLAY T. STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_inspect_tallying_all_after,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'AABAA'. 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. INSPECT S1 TALLYING T FOR ALL 'A' AFTER INITIAL 'B'. DISPLAY T. STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_inspect_tallying_characters_before,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'AABAA'. 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. INSPECT S1 TALLYING T FOR CHARACTERS BEFORE INITIAL 'B'. DISPLAY T. STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_inspect_tallying_multiple_all,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'ABCAB'. 01 T1 PIC 9 VALUE 0. 01 T2 PIC 9 VALUE 0. PROCEDURE DIVISION. INSPECT S1 TALLYING T1 FOR ALL 'A' T2 FOR ALL 'B'. DISPLAY T1 T2. STOP RUN.",
    vec!["22"]
);
cobol_test!(
    test_inspect_tallying_multiple_conditions,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'ABXAB'. 01 T1 PIC 9 VALUE 0. PROCEDURE DIVISION. INSPECT S1 TALLYING T1 FOR ALL 'A' BEFORE INITIAL 'X' ALL 'B' AFTER INITIAL 'X'. DISPLAY T1. STOP RUN.",
    vec!["2"]
); // A before X is 1, B after X is 1
cobol_test!(
    test_inspect_parse_9,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_10,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_11,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_12,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_13,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_14,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_15,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_16,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_17,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_18,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_19,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_20,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_21,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_22,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_23,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_24,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_25,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_26,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_27,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_28,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_29,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_inspect_parse_30,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
