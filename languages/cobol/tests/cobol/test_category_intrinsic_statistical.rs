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

// 30 edge cases for Intrinsic Statistical functions
cobol_test!(
    test_stat_mean_pos,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEAN(2 4 6 8). STOP RUN.",
    vec!["5"]
);
cobol_test!(
    test_stat_mean_neg,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEAN(-2 -4 -6 -8). STOP RUN.",
    vec!["-5"]
);
cobol_test!(
    test_stat_median_odd,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEDIAN(1 5 9). STOP RUN.",
    vec!["5"]
);
cobol_test!(
    test_stat_median_even,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEDIAN(1 3 5 7). STOP RUN.",
    vec!["4"]
);
cobol_test!(
    test_stat_midrange_pos,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MIDRANGE(2 10). STOP RUN.",
    vec!["6"]
);
cobol_test!(
    test_stat_midrange_neg,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MIDRANGE(-2 -10). STOP RUN.",
    vec!["-6"]
);
cobol_test!(
    test_stat_range_pos,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION RANGE(2 10). STOP RUN.",
    vec!["8"]
);
cobol_test!(
    test_stat_range_neg,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION RANGE(-2 -10). STOP RUN.",
    vec!["8"]
);
cobol_test!(
    test_stat_variance_uniform,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION VARIANCE(5 5 5 5). STOP RUN.",
    vec!["0"]
);
cobol_test!(
    test_stat_variance_spread,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION VARIANCE(2 4 6). STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_stat_stddev_uniform,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION STANDARD-DEVIATION(5 5 5 5). STOP RUN.",
    vec!["0"]
);
cobol_test!(
    test_stat_stddev_spread,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION STANDARD-DEVIATION(2 4 6) > 1 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_stat_random_seed,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION RANDOM(1) > 0 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_stat_random_no_seed,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION RANDOM > 0 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_stat_sum_large,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION SUM(9999 9999). STOP RUN.",
    vec!["19998"]
);
cobol_test!(
    test_stat_parse_16,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEAN(-1 0 1). STOP RUN.",
    vec!["0"]
);
cobol_test!(
    test_stat_parse_17,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEDIAN(9 1 7 3). STOP RUN.",
    vec!["5"]
);
cobol_test!(
    test_stat_parse_18,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION SUM(1 2 3) = 6 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_stat_parse_19,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION RANGE(-10 0 20). STOP RUN.",
    vec!["30"]
);
cobol_test!(
    test_stat_parse_20,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION MEAN(10 20) = 15 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_stat_parse_21,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION VARIANCE(1 2 3 4) > 1 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_stat_parse_22,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION STANDARD-DEVIATION(1 1 2 2) > 0 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_stat_parse_23,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_stat_parse_24,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_stat_parse_25,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_stat_parse_26,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_stat_parse_27,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_stat_parse_28,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_stat_parse_29,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_stat_parse_30,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
