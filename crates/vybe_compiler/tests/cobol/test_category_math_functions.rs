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

// 30 distinct tests for Math Functions
cobol_test!(
    test_math_fn_sum,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION SUM(1 2 3). STOP RUN.",
    vec!["6"]
);
cobol_test!(
    test_math_fn_mean,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEAN(2 4 6). STOP RUN.",
    vec!["4"]
);
cobol_test!(
    test_math_fn_median,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEDIAN(1 9 5). STOP RUN.",
    vec!["5"]
);
cobol_test!(
    test_math_fn_median_even,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEDIAN(1 5 9 13). STOP RUN.",
    vec!["7"]
);
cobol_test!(
    test_math_fn_midrange,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MIDRANGE(1 9). STOP RUN.",
    vec!["5"]
);
cobol_test!(
    test_math_fn_range,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION RANGE(1 9). STOP RUN.",
    vec!["8"]
);
cobol_test!(
    test_math_fn_variance,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION VARIANCE(2 4 6). STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_math_fn_standard_deviation,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION STANDARD-DEVIATION(2 4 6) > 1 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_math_fn_mod,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MOD(10 3). STOP RUN.",
    vec!["1"]
);
cobol_test!(
    test_math_fn_rem,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION REM(10 3). STOP RUN.",
    vec!["1"]
);
cobol_test!(
    test_math_fn_integer,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION INTEGER(2.5). STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_math_fn_integer_part,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION INTEGER-PART(2.5). STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_math_fn_abs_pos,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ABS(5). STOP RUN.",
    vec!["5"]
);
cobol_test!(
    test_math_fn_abs_neg,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ABS(-5). STOP RUN.",
    vec!["5"]
);
cobol_test!(
    test_math_fn_sign_pos,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION SIGN(5). STOP RUN.",
    vec!["1"]
);
cobol_test!(
    test_math_fn_sign_neg,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION SIGN(-5). STOP RUN.",
    vec!["-1"]
);
cobol_test!(
    test_math_fn_sign_zero,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION SIGN(0). STOP RUN.",
    vec!["0"]
);
cobol_test!(
    test_math_fn_sqrt,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION SQRT(16). STOP RUN.",
    vec!["4"]
);
cobol_test!(
    test_math_fn_log,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION LOG(10) > 2 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_math_fn_log10,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION LOG10(100). STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_math_fn_exp,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION EXP(1) > 2.7 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_math_fn_exp10,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION EXP10(2). STOP RUN.",
    vec!["100"]
);
cobol_test!(
    test_math_fn_sin,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION SIN(0) = 0 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_math_fn_cos,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION COS(0) = 1 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_math_fn_tan,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION TAN(0) = 0 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_math_fn_asin,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION ASIN(0) = 0 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_math_fn_acos,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION ACOS(1) = 0 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_math_fn_atan,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION ATAN(0) = 0 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_math_fn_pi,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION PI > 3.14 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_math_fn_e,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION E > 2.71 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
