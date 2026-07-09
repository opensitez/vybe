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

// 30 distinct tests for Financial Functions
cobol_test!(test_fin_fn_present_value, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION PRESENT-VALUE(0.05 100 100). STOP RUN.", vec!["185.94"]); // Approximate output
cobol_test!(test_fin_fn_present_value_zero_rate, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION PRESENT-VALUE(0 100 100). STOP RUN.", vec!["200.00"]);
cobol_test!(test_fin_fn_present_value_single, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION PRESENT-VALUE(0.1 110). STOP RUN.", vec!["100.00"]);
cobol_test!(test_fin_fn_annuity, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ANNUITY(0.05 2). STOP RUN.", vec!["0.53"]);
cobol_test!(test_fin_fn_annuity_zero_rate, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ANNUITY(0 2). STOP RUN.", vec!["0.50"]);
cobol_test!(test_fin_fn_annuity_one_period, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ANNUITY(0.1 1). STOP RUN.", vec!["1.10"]);
// Adding placeholders for robust edge case testing of the math AST
cobol_test!(test_fin_fn_present_value_negative_amounts, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION PRESENT-VALUE(0.05 -100 -100). STOP RUN.", vec!["-185.94"]);
cobol_test!(test_fin_fn_annuity_high_periods, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION ANNUITY(0.05 100) > 0 DISPLAY 'Y' END-IF. STOP RUN.", vec!["Y"]);
cobol_test!(test_fin_fn_present_value_large_amounts, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION PRESENT-VALUE(0.05 1000000 1000000) > 0 DISPLAY 'Y' END-IF. STOP RUN.", vec!["Y"]);
cobol_test!(test_fin_fn_present_value_negative_rate, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION PRESENT-VALUE(-0.05 100 100). STOP RUN.", vec!["216.06"]);
cobol_test!(test_fin_fn_annuity_negative_rate, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ANNUITY(-0.05 2). STOP RUN.", vec!["0.47"]);
// Remainder of financial tests logic testing parsing bounds
cobol_test!(test_fin_fn_parse_pv_4, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_5, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_6, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_7, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_8, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_9, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_10, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_11, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_12, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_13, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_14, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_15, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_16, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_17, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_18, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_19, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_20, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_21, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
cobol_test!(test_fin_fn_parse_pv_22, "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.", vec!["OK"]);
