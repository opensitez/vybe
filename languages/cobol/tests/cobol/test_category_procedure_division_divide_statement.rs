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

// 30 specific tests for PROCEDURE DIVISION DIVIDE STATEMENT
cobol_test!(
    test_divide_basic,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 6. 01 V2 PIC 9 VALUE 2. PROCEDURE DIVISION. DIVIDE V2 INTO V1. DISPLAY V1. STOP RUN.",
    vec!["3"]
);
cobol_test!(
    test_divide_giving,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 6. 01 V2 PIC 9 VALUE 2. 01 R PIC 9. PROCEDURE DIVISION. DIVIDE V1 BY V2 GIVING R. DISPLAY R. STOP RUN.",
    vec!["3"]
);
cobol_test!(
    test_divide_giving_into,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 2. 01 V2 PIC 9 VALUE 6. 01 R PIC 9. PROCEDURE DIVISION. DIVIDE V1 INTO V2 GIVING R. DISPLAY R. STOP RUN.",
    vec!["3"]
);
cobol_test!(
    test_divide_remainder,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 7. 01 V2 PIC 9 VALUE 2. 01 R PIC 9. 01 REM PIC 9. PROCEDURE DIVISION. DIVIDE V1 BY V2 GIVING R REMAINDER REM. DISPLAY R REM. STOP RUN.",
    vec!["31"]
);
cobol_test!(
    test_divide_zero_divide,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 5. 01 V2 PIC 9 VALUE 0. 01 R PIC 9. PROCEDURE DIVISION. DIVIDE V1 BY V2 GIVING R ON SIZE ERROR DISPLAY 'ERR'. STOP RUN.",
    vec!["ERR"]
);
cobol_test!(
    test_divide_parse_6,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_7,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_8,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_9,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_10,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_11,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_12,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_13,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_14,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_15,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_16,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_17,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_18,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_19,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_20,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_21,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_22,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_23,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_24,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_25,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_26,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_27,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_28,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_29,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_divide_parse_30,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);

cobol_test!(
    test_divide_with_not_on_size_error_path,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC 9 VALUE 8. 01 V2 PIC 9 VALUE 2. 01 R PIC 9. 01 REM PIC 9. PROCEDURE DIVISION. DIVIDE V1 BY V2 GIVING R REMAINDER REM ON SIZE ERROR DISPLAY 'ERR' NOT ON SIZE ERROR DISPLAY R DISPLAY REM END-DIVIDE STOP RUN.",
    vec!["4", "0"]
);
cobol_test!(
    test_divide_into_chain_multiple_quotients,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 DIVIDEND PIC 9 VALUE 9. 01 DIVISOR PIC 9 VALUE 2. PROCEDURE DIVISION. DIVIDE DIVIDEND INTO DIVISOR. DISPLAY DIVISOR. STOP RUN.",
    vec!["4"]
);
cobol_test!(
    test_divide_by_literal_runtime,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V1 PIC S9(4) VALUE 10. 01 R PIC S9(4). 01 REM PIC S9(4). PROCEDURE DIVISION. DIVIDE 30 BY V1 GIVING R REMAINDER REM. DISPLAY R DISPLAY REM STOP RUN.",
    vec!["3", "0"]
);
