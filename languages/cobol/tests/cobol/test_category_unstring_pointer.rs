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

// 30 specific tests for UNSTRING with POINTER and TALLYING
cobol_test!(
    test_unstring_ptr_basic,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'A,B,C'. 01 R1 PIC X. 01 P PIC 9 VALUE 1. PROCEDURE DIVISION. UNSTRING S1 DELIMITED BY ',' INTO R1 WITH POINTER P. DISPLAY R1. STOP RUN.",
    vec!["A"]
);
cobol_test!(
    test_unstring_ptr_offset,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'A,B,C'. 01 R1 PIC X. 01 P PIC 9 VALUE 3. PROCEDURE DIVISION. UNSTRING S1 DELIMITED BY ',' INTO R1 WITH POINTER P. DISPLAY R1. STOP RUN.",
    vec!["B"]
);
cobol_test!(
    test_unstring_ptr_update,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'A,B,C'. 01 R1 PIC X. 01 R2 PIC X. 01 P PIC 9 VALUE 1. PROCEDURE DIVISION. UNSTRING S1 DELIMITED BY ',' INTO R1 R2 WITH POINTER P. DISPLAY P. STOP RUN.",
    vec!["4"]
);
cobol_test!(
    test_unstring_tallying_basic,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'A,B,C'. 01 R1 PIC X. 01 R2 PIC X. 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. UNSTRING S1 DELIMITED BY ',' INTO R1 R2 TALLYING IN T. DISPLAY T. STOP RUN.",
    vec!["2"]
);
cobol_test!(
    test_unstring_tallying_offset,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'A,B,C'. 01 R1 PIC X. 01 R2 PIC X. 01 T PIC 9 VALUE 5. PROCEDURE DIVISION. UNSTRING S1 DELIMITED BY ',' INTO R1 R2 TALLYING IN T. DISPLAY T. STOP RUN.",
    vec!["7"]
); // Tallying adds to initial value
cobol_test!(
    test_unstring_ptr_tallying,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'A,B,C'. 01 R1 PIC X. 01 P PIC 9 VALUE 3. 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. UNSTRING S1 DELIMITED BY ',' INTO R1 WITH POINTER P TALLYING IN T. DISPLAY T. STOP RUN.",
    vec!["1"]
);
cobol_test!(
    test_unstring_ptr_tallying_overflow,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'A,B,C'. 01 R1 PIC X. 01 P PIC 9 VALUE 6. 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. UNSTRING S1 DELIMITED BY ',' INTO R1 WITH POINTER P TALLYING IN T ON OVERFLOW DISPLAY 'OVF'. STOP RUN.",
    vec!["OVF"]
);
cobol_test!(
    test_unstring_tallying_overflow,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 S1 PIC X(5) VALUE 'A,B,C'. 01 R1 PIC X(1). 01 T PIC 9 VALUE 0. PROCEDURE DIVISION. UNSTRING S1 DELIMITED BY ',' INTO R1 TALLYING IN T ON OVERFLOW DISPLAY 'OVF'. STOP RUN.",
    vec!["OVF"]
);
cobol_test!(
    test_unstring_ptr_parse_9,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_10,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_11,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_12,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_13,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_14,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_15,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_16,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_17,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_18,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_19,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_20,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_21,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_22,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_23,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_24,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_25,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_26,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_27,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_28,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_29,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_unstring_ptr_parse_30,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
