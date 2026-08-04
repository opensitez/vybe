*> vybe-test: cobol/currency_sign/currency_sign_with_computation_target_compiles
*> origin: languages/cobol/tests/cobol/test_currency_sign.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CUR9.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    CURRENCY SIGN IS "$".
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3)V99 VALUE 10.00.
01 B PIC 9(3)V99 VALUE 5.00.
01 DST PIC $ZZ9.99.
PROCEDURE DIVISION.
    ADD A B GIVING A.
    MOVE A TO DST.
    STOP RUN.

