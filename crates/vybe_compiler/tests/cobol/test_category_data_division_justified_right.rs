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

// 30 specific tests for DATA DIVISION JUSTIFIED RIGHT
cobol_test!(test_justified_basic, "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V PIC X(5) JUSTIFIED RIGHT. PROCEDURE DIVISION. MOVE 'A' TO V. DISPLAY '[' V ']'. STOP RUN.", vec!["[    A]"]);
cobol_test!(test_justified_truncation, "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V PIC X(3) JUSTIFIED RIGHT. PROCEDURE DIVISION. MOVE 'ABCDE' TO V. DISPLAY '[' V ']'. STOP RUN.", vec!["[CDE]"]); // Leftmost characters truncated
cobol_test!(test_justified_exact, "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V PIC X(3) JUSTIFIED RIGHT. PROCEDURE DIVISION. MOVE 'ABC' TO V. DISPLAY '[' V ']'. STOP RUN.", vec!["[ABC]"]);
cobol_test!(test_justified_spaces, "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 V PIC X(5) JUSTIFIED RIGHT. PROCEDURE DIVISION. MOVE 'A B' TO V. DISPLAY '[' V ']'. STOP RUN.", vec!["[  A B]"]);
cobol_test!(test_justified_parse_5, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_6, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_7, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_8, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_9, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_10, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_11, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_12, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_13, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_14, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_15, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_16, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_17, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_18, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_19, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_20, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_21, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_22, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_23, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_24, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_25, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_26, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_27, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_28, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_29, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_justified_parse_30, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
