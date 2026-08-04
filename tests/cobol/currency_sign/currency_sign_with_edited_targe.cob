*> vybe-test: cobol/currency_sign/currency_sign_with_edited_target_compiles
*> origin: languages/cobol/tests/cobol/test_currency_sign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CUR6.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CURRENCY SIGN IS "$".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC 9(4)V99 VALUE 9999.99.
01 DST PIC $,$$9.99.
PROCEDURE DIVISION.
    MOVE SRC TO DST.
    STOP RUN.

