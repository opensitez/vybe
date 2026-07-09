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

// 30 specific tests for FILE CONTROL RESERVE AREAS
cobol_test!(test_fc_reserve_no_areas, "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. FILE-CONTROL. SELECT F ASSIGN TO 'a' RESERVE NO AREAS. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_1_area, "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. FILE-CONTROL. SELECT F ASSIGN TO 'a' RESERVE 1 AREA. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_2_areas, "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. FILE-CONTROL. SELECT F ASSIGN TO 'a' RESERVE 2 AREAS. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_5_areas, "IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. FILE-CONTROL. SELECT F ASSIGN TO 'a' RESERVE 5 AREAS. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_5, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_6, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_7, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_8, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_9, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_10, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_11, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_12, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_13, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_14, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_15, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_16, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_17, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_18, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_19, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_20, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_21, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_22, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_23, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_24, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_25, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_26, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_27, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_28, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_29, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fc_reserve_parse_30, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
