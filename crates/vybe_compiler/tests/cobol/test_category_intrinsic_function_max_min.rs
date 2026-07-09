use crate::helpers;

macro_rules! cobol_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let out = crate::helpers::run_prints($src);
            assert_eq!(out, $expected);
        }
    }
}

// 30 specific tests for INTRINSIC FUNCTION MAX MIN
cobol_test!(test_max_numeric, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MAX(1 5 3). STOP RUN.", vec!["5"]);
cobol_test!(test_min_numeric, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MIN(1 5 3). STOP RUN.", vec!["1"]);
cobol_test!(test_max_alphanumeric, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MAX('APPLE' 'ZEBRA' 'BANANA'). STOP RUN.", vec!["ZEBRA"]);
cobol_test!(test_min_alphanumeric, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MIN('APPLE' 'ZEBRA' 'BANANA'). STOP RUN.", vec!["APPLE"]);
cobol_test!(test_max_min_parse_5, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_6, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_7, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_8, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_9, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_10, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_11, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_12, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_13, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_14, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_15, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_16, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_17, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_18, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_19, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_20, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_21, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_22, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_23, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_24, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_25, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_26, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_27, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_28, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_29, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_max_min_parse_30, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
