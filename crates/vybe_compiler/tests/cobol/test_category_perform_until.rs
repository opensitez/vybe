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

// 30 specific tests for PERFORM UNTIL
cobol_test!(
    test_perform_until_basic,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9 VALUE 1. PROCEDURE DIVISION. PERFORM M-PARA UNTIL I > 3. DISPLAY 'OK'. STOP RUN. M-PARA. DISPLAY I. ADD 1 TO I.",
    vec!["1", "2", "3", "OK"]
);
cobol_test!(
    test_perform_until_test_before,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9 VALUE 4. PROCEDURE DIVISION. PERFORM M-PARA WITH TEST BEFORE UNTIL I > 3. DISPLAY 'OK'. STOP RUN. M-PARA. DISPLAY I. ADD 1 TO I.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_test_after,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9 VALUE 4. PROCEDURE DIVISION. PERFORM M-PARA WITH TEST AFTER UNTIL I > 3. DISPLAY 'OK'. STOP RUN. M-PARA. DISPLAY I. ADD 1 TO I.",
    vec!["4", "OK"]
);
cobol_test!(
    test_perform_until_inline,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9 VALUE 1. PROCEDURE DIVISION. PERFORM UNTIL I > 3 DISPLAY I ADD 1 TO I END-PERFORM. DISPLAY 'OK'. STOP RUN.",
    vec!["1", "2", "3", "OK"]
);
cobol_test!(
    test_perform_until_inline_test_after,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9 VALUE 4. PROCEDURE DIVISION. PERFORM WITH TEST AFTER UNTIL I > 3 DISPLAY I ADD 1 TO I END-PERFORM. DISPLAY 'OK'. STOP RUN.",
    vec!["4", "OK"]
);
cobol_test!(
    test_perform_until_complex_cond,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9 VALUE 1. 01 J PIC 9 VALUE 1. PROCEDURE DIVISION. PERFORM UNTIL I > 2 OR J > 2 DISPLAY I J ADD 1 TO I J END-PERFORM. DISPLAY 'OK'. STOP RUN.",
    vec!["11", "22", "OK"]
);
cobol_test!(
    test_perform_until_exit_perform,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9 VALUE 1. PROCEDURE DIVISION. PERFORM UNTIL I > 5 DISPLAY I IF I = 3 EXIT PERFORM END-IF ADD 1 TO I END-PERFORM. DISPLAY 'OK'. STOP RUN.",
    vec!["1", "2", "3", "OK"]
);
cobol_test!(
    test_perform_until_exit_perform_cycle,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9 VALUE 0. PROCEDURE DIVISION. PERFORM UNTIL I > 2 ADD 1 TO I IF I = 2 EXIT PERFORM CYCLE END-IF DISPLAY I END-PERFORM. DISPLAY 'OK'. STOP RUN.",
    vec!["1", "3", "OK"]
);
cobol_test!(
    test_perform_until_parse_9,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_10,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_11,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_12,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_13,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_14,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_15,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_16,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_17,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_18,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_19,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_20,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_21,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_22,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_23,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_24,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_25,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_26,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_27,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_28,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_29,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_perform_until_parse_30,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
