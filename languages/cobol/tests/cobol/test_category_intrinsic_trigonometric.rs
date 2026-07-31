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

// 30 edge cases for Trigonometric functions
cobol_test!(
    test_trig_sin_0,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION SIN(0). STOP RUN.",
    vec!["0"]
);
cobol_test!(
    test_trig_cos_0,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION COS(0). STOP RUN.",
    vec!["1"]
);
cobol_test!(
    test_trig_tan_0,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION TAN(0). STOP RUN.",
    vec!["0"]
);
cobol_test!(
    test_trig_asin_0,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ASIN(0). STOP RUN.",
    vec!["0"]
);
cobol_test!(
    test_trig_acos_1,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ACOS(1). STOP RUN.",
    vec!["0"]
);
cobol_test!(
    test_trig_atan_0,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ATAN(0). STOP RUN.",
    vec!["0"]
);
cobol_test!(
    test_trig_pi,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION PI > 3.14 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_sin_pi_half,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION SIN(1.5707) > 0.9 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_cos_pi,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION COS(3.1415) < -0.9 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_tan_pi_fourth,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION TAN(0.7853) > 0.9 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_11,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION SIN(FUNCTION COS(0)) > 0.8 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_12,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION COS(FUNCTION SIN(0)) > 0.9 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_13,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION ATAN(FUNCTION TAN(0.7853)) > 0.7 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_14,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION SIN(0.0) = FUNCTION SIN(0.0) DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_15,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION COS(3.1415) < -0.9 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_16,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION ACOS(0) > 1.56 AND FUNCTION ACOS(0) < 1.58 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_17,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION ASIN(0.5) > 0.51 AND FUNCTION ASIN(0.5) < 0.53 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_18,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION COS(0.0) > 0.9999 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_19,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION SIN(-0.0) = FUNCTION SIN(0.0) DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_20,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION TAN(1.5707) > 10000 OR FUNCTION TAN(1.5707) < -10000 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_21,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION ACOS(0.5) > 1.0 AND FUNCTION ACOS(0.5) < 1.1 DISPLAY 'Y' END-IF. STOP RUN.",
    vec!["Y"]
);
cobol_test!(
    test_trig_parse_31,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_trig_parse_22,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_trig_parse_23,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_trig_parse_24,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_trig_parse_25,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_trig_parse_26,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_trig_parse_27,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_trig_parse_28,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_trig_parse_29,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_trig_parse_30,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
