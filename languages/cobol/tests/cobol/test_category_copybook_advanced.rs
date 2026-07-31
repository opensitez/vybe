use crate::helpers;

macro_rules! cobol_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let out = crate::helpers::run_prints($src);
            let expected: Vec<String> = $expected.into_iter().map(|s: &str| s.to_string()).collect();
            assert_eq!(out, expected);
        }
    };
}

cobol_test!(
    test_cp_basic,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. COPY 'MOCK'. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
); // Parse test
cobol_test!(
    test_cp_replacing,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. COPY 'MOCK' REPLACING ==A== BY ==B==. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_cp_replacing_multiple,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. COPY 'MOCK' REPLACING ==A== BY ==B== ==C== BY ==D==. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_cp_replacing_word,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. COPY 'MOCK' REPLACING WORD1 BY WORD2. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_cp_replacing_string,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. COPY 'MOCK' REPLACING 'A' BY 'B'. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_cp_replacing_identifier,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. COPY 'MOCK' REPLACING ID-1 BY ID-2. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_cp_replacing_chained_identifiers,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. COPY 'MOCK' REPLACING ALIAS-A BY ALIAS-B ALIAS-B BY ALIAS-C. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_cp_suppress,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. COPY 'MOCK' SUPPRESS. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_cp_replace_in_data_division,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. REPLACE ==FIELD-ONE== BY ==FIELD-TWO==. 01 F PIC X VALUE 'Z'. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_cp_replace_off_scoped,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. REPLACE ==A== BY ==B==. REPLACE OFF. 01 A PIC X VALUE '1'. PROCEDURE DIVISION. DISPLAY A. STOP RUN.",
    vec!["1"]
);
cobol_test!(
    test_cp_in_library,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. COPY 'MOCK' IN 'LIB'. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_cp_in_library_with_replacing,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. COPY 'MOCK' IN 'LIB' REPLACING ==A== BY ==B==. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_cp_copy_in_linkage_section,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. LINKAGE SECTION. COPY 'MOCK'. PROCEDURE DIVISION. DISPLAY 'OK'. STOP RUN.",
    vec!["OK"]
);
cobol_test!(
    test_replace_statement,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. REPLACE ==A== BY ==B==. 01 A PIC X VALUE '1'. PROCEDURE DIVISION. DISPLAY B. STOP RUN.",
    vec!["1"]
);
cobol_test!(
    test_replace_off,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. REPLACE ==A== BY ==B==. REPLACE OFF. 01 A PIC X VALUE '1'. PROCEDURE DIVISION. DISPLAY A. STOP RUN.",
    vec!["1"]
);
