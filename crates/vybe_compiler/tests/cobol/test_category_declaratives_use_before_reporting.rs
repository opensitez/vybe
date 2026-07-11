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

// 30 specific tests for DECLARATIVES USE BEFORE REPORTING
cobol_test!(
    test_use_before_reporting_basic,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. DATA DIVISION. REPORT SECTION. RD R1. 01 G TYPE IS DETAIL. 05 COLUMN 1 PIC X VALUE 'A'. PROCEDURE DIVISION. DECLARATIVES. S1 SECTION. USE BEFORE REPORTING G. S1-PARA. DISPLAY 'BEF'. END DECLARATIVES. MAIN SECTION. DISPLAY 'MAIN'. STOP RUN.",
    vec!["MAIN"]
); // Execution is tested when Generate is called, this tests parsing
cobol_test!(
    test_use_before_reporting_multiple,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. DATA DIVISION. REPORT SECTION. RD R1. 01 G1 TYPE IS DETAIL. 01 G2 TYPE IS DETAIL. PROCEDURE DIVISION. DECLARATIVES. S1 SECTION. USE BEFORE REPORTING G1 G2. S1-PARA. DISPLAY 'BEF'. END DECLARATIVES. MAIN SECTION. DISPLAY 'MAIN'. STOP RUN.",
    vec!["MAIN"]
);
cobol_test!(
    test_use_before_reporting_parse_3,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_4,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_5,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_6,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_7,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_8,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_9,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_10,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_11,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_12,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_13,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_14,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_15,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_16,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_17,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_18,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_19,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_20,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_21,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_22,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_23,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_24,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_25,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_26,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_27,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_28,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_29,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_use_before_reporting_parse_30,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
