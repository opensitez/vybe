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

cobol_test!(
    test_dd_value_truncation_right,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 R PIC X(2) VALUE 'ABC'. PROCEDURE DIVISION. DISPLAY R. STOP RUN.",
    vec!["AB"]
);
cobol_test!(
    test_dd_value_truncation_left,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 R PIC 9(2) VALUE 123. PROCEDURE DIVISION. DISPLAY R. STOP RUN.",
    vec!["23"]
);
cobol_test!(
    test_dd_value_pad_right,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 R PIC X(4) VALUE 'A'. PROCEDURE DIVISION. DISPLAY '[' R ']'. STOP RUN.",
    vec!["[A   ]"]
);
cobol_test!(
    test_dd_value_pad_left,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 R PIC 9(4) VALUE 1. PROCEDURE DIVISION. DISPLAY '[' R ']'. STOP RUN.",
    vec!["[0001]"]
);
cobol_test!(
    test_dd_redefines_smaller,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC X(4) VALUE '1234'. 01 B REDEFINES A PIC X(2). PROCEDURE DIVISION. DISPLAY B. STOP RUN.",
    vec!["12"]
);
cobol_test!(
    test_dd_redefines_larger_error,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC X(2) VALUE '12'. 01 B REDEFINES A PIC X(4). PROCEDURE DIVISION. DISPLAY B. STOP RUN.",
    vec!["12"]
); // Might just pad or warning depending on compiler
cobol_test!(
    test_dd_redefines_nested,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 A PIC X(4) VALUE '1234'. 01 B REDEFINES A. 05 B1 PIC X(2). 05 B2 PIC X(2). PROCEDURE DIVISION. DISPLAY B2. STOP RUN.",
    vec!["34"]
);
cobol_test!(
    test_dd_renames_single,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 G. 05 A PIC X VALUE '1'. 05 B PIC X VALUE '2'. 66 R RENAMES A. PROCEDURE DIVISION. DISPLAY R. STOP RUN.",
    vec!["1"]
);
cobol_test!(
    test_dd_renames_thru,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 G. 05 A PIC X VALUE '1'. 05 B PIC X VALUE '2'. 05 C PIC X VALUE '3'. 66 R RENAMES A THRU B. PROCEDURE DIVISION. DISPLAY R. STOP RUN.",
    vec!["12"]
);
cobol_test!(
    test_dd_blank_when_zero_group,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 G BLANK WHEN ZERO. 05 A PIC 9(2) VALUE 0. PROCEDURE DIVISION. DISPLAY '[' G ']'. STOP RUN.",
    vec!["[  ]"]
); // Only numeric or numeric edited can have BLANK WHEN ZERO, but some compilers allow on group. Parse test.
