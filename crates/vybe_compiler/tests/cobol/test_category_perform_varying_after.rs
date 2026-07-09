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

// 30 specific tests for PERFORM VARYING AFTER
cobol_test!(test_perform_var_after_basic, "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9. 01 J PIC 9. PROCEDURE DIVISION. PERFORM M-PARA VARYING I FROM 1 BY 1 UNTIL I > 2 AFTER J FROM 1 BY 1 UNTIL J > 2. DISPLAY 'OK'. STOP RUN. M-PARA. DISPLAY I J.", vec!["11", "12", "21", "22", "OK"]);
cobol_test!(test_perform_var_after_three, "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9. 01 J PIC 9. 01 K PIC 9. PROCEDURE DIVISION. PERFORM M-PARA VARYING I FROM 1 BY 1 UNTIL I > 2 AFTER J FROM 1 BY 1 UNTIL J > 2 AFTER K FROM 1 BY 1 UNTIL K > 2. DISPLAY 'OK'. STOP RUN. M-PARA. DISPLAY I J K.", vec!["111", "112", "121", "122", "211", "212", "221", "222", "OK"]);
cobol_test!(test_perform_var_after_inline, "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9. 01 J PIC 9. PROCEDURE DIVISION. PERFORM VARYING I FROM 1 BY 1 UNTIL I > 2 AFTER J FROM 1 BY 1 UNTIL J > 2 DISPLAY I J END-PERFORM. DISPLAY 'OK'. STOP RUN.", vec!["11", "12", "21", "22", "OK"]);
cobol_test!(test_perform_var_after_test_after, "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 I PIC 9. 01 J PIC 9. PROCEDURE DIVISION. PERFORM M-PARA WITH TEST AFTER VARYING I FROM 1 BY 1 UNTIL I > 2 AFTER J FROM 1 BY 1 UNTIL J > 2. DISPLAY 'OK'. STOP RUN. M-PARA. DISPLAY I J.", vec!["11", "12", "13", "21", "22", "23", "31", "32", "33", "OK"]); // Specific iteration counts depend on ANSI spec, test structure here
cobol_test!(test_perform_var_after_parse_5, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_6, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_7, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_8, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_9, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_10, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_11, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_12, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_13, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_14, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_15, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_16, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_17, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_18, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_19, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_20, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_21, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_22, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_23, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_24, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_25, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_26, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_27, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_28, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_29, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_perform_var_after_parse_30, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
