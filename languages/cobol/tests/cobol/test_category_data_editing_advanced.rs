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
    test_edit_z_suppress_all_zeros,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC ZZZ. 01 Y PIC 999 VALUE 0. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[   ]"]
);
cobol_test!(
    test_edit_z_suppress_partial_zeros,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC ZZ9. 01 Y PIC 999 VALUE 0. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[  0]"]
);
cobol_test!(
    test_edit_asterisk_fill_all_zeros,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC ***. 01 Y PIC 999 VALUE 0. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[***]"]
);
cobol_test!(
    test_edit_asterisk_fill_partial_zeros,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC **9. 01 Y PIC 999 VALUE 0. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[**0]"]
);
cobol_test!(
    test_edit_minus_sign_positive,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC ---. 01 Y PIC S999 VALUE 5. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[  5]"]
);
cobol_test!(
    test_edit_minus_sign_negative,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC ---. 01 Y PIC S999 VALUE -5. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[ -5]"]
);
cobol_test!(
    test_edit_plus_sign_positive,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC +++. 01 Y PIC S999 VALUE 5. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[ +5]"]
);
cobol_test!(
    test_edit_plus_sign_negative,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC +++. 01 Y PIC S999 VALUE -5. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[ -5]"]
);
cobol_test!(
    test_edit_cr_positive,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC 99CR. 01 Y PIC S99 VALUE 5. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[05  ]"]
);
cobol_test!(
    test_edit_cr_negative,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC 99CR. 01 Y PIC S99 VALUE -5. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[05CR]"]
);
cobol_test!(
    test_edit_db_positive,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC 99DB. 01 Y PIC S99 VALUE 5. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[05  ]"]
);
cobol_test!(
    test_edit_db_negative,
    "IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 X PIC 99DB. 01 Y PIC S99 VALUE -5. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'. STOP RUN.",
    vec!["[05DB]"]
);
#[test]
fn compile_edit_currency_symbol_clause() {
    helpers::compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC $$,999.\n01 Y PIC 999 VALUE 12.\nPROCEDURE DIVISION.\n    MOVE Y TO X\n    STOP RUN.",
    );
}
#[test]
fn compile_edit_picture_v_significant_zeros() {
    helpers::compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC ZZ,ZZ9.\n01 Y PIC 9999 VALUE 100.\nPROCEDURE DIVISION.\n    MOVE Y TO X.\n    STOP RUN.",
    );
}
#[test]
fn compile_edit_currency_symbol_and_comma() {
    helpers::compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC $$,$$$.99.\n01 Y PIC 9999V99 VALUE 123456.\nPROCEDURE DIVISION.\n    MOVE Y TO X.\n    STOP RUN.",
    );
}
#[test]
fn compile_edit_alphanumeric_editing_clauses() {
    helpers::compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC X(10).\n01 Y PIC 9(4) VALUE 100.\nPROCEDURE DIVISION.\n    MOVE Y TO X.\n    STOP RUN.",
    );
}
#[test]
fn compile_edit_redefines_with_edit_pictures() {
    helpers::compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01  RAW PIC 9(4) VALUE 1234.\n01 OUT REDEFINES RAW PIC ZZ,ZZ.\nPROCEDURE DIVISION.\n    DISPLAY OUT.\n    STOP RUN.",
    );
}
#[test]
fn compile_edit_trailing_sign_with_currency() {
    helpers::compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC ZZ9CR.\n01 Y PIC S999 VALUE -12.\nPROCEDURE DIVISION.\n    MOVE Y TO X.\n    STOP RUN.",
    );
}
