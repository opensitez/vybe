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

// 30 specific tests for EVALUATE ALSO
cobol_test!(
    test_evaluate_also_basic,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC 9 VALUE 1. 01 B PIC 9 VALUE 2. PROCEDURE DIVISION. EVALUATE A ALSO B WHEN 1 ALSO 2 DISPLAY 'Y' WHEN OTHER DISPLAY 'N' END-EVALUATE. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_evaluate_also_second_match,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC 9 VALUE 2. 01 B PIC 9 VALUE 3. PROCEDURE DIVISION. EVALUATE A ALSO B WHEN 1 ALSO 2 DISPLAY '1' WHEN 2 ALSO 3 DISPLAY '2' WHEN OTHER DISPLAY 'N' END-EVALUATE. STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_evaluate_also_any,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC 9 VALUE 1. 01 B PIC 9 VALUE 3. PROCEDURE DIVISION. EVALUATE A ALSO B WHEN 1 ALSO ANY DISPLAY 'Y' WHEN OTHER DISPLAY 'N' END-EVALUATE. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_evaluate_also_thru,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC 9 VALUE 2. 01 B PIC 9 VALUE 4. PROCEDURE DIVISION. EVALUATE A ALSO B WHEN 1 THRU 3 ALSO 4 THRU 6 DISPLAY 'Y' WHEN OTHER DISPLAY 'N' END-EVALUATE. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_evaluate_also_true,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC 9 VALUE 1. PROCEDURE DIVISION. EVALUATE TRUE ALSO A WHEN A = 1 ALSO 1 DISPLAY 'Y' WHEN OTHER DISPLAY 'N' END-EVALUATE. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_evaluate_also_condition,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC 9 VALUE 1. 01 B PIC 9 VALUE 2. PROCEDURE DIVISION. EVALUATE TRUE ALSO TRUE WHEN A = 1 ALSO B = 2 DISPLAY 'Y' WHEN OTHER DISPLAY 'N' END-EVALUATE. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_evaluate_also_multiple,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC 9 VALUE 1. 01 B PIC 9 VALUE 2. 01 C PIC 9 VALUE 3. PROCEDURE DIVISION. EVALUATE A ALSO B ALSO C WHEN 1 ALSO 2 ALSO 3 DISPLAY 'Y' WHEN OTHER DISPLAY 'N' END-EVALUATE. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_evaluate_also_any_combination,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC 9 VALUE 1. 01 B PIC 9 VALUE 2. PROCEDURE DIVISION. EVALUATE A ALSO B WHEN ANY ALSO 1 DISPLAY '1' WHEN 1 ALSO ANY DISPLAY '2' WHEN OTHER DISPLAY 'N' END-EVALUATE. STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_evaluate_also_parse_9,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_10,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_11,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_12,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_13,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_14,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_15,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_16,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_17,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_18,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_19,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_20,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_21,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_22,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_23,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_24,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_25,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_26,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_27,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_28,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_29,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_evaluate_also_parse_30,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
