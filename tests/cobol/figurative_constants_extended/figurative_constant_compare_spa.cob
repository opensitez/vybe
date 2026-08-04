*> vybe-test: cobol/figurative_constants_extended/figurative_constant_compare_spaces_and_zeros
*> origin: languages/cobol/tests/cobol/test_figurative_constants_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(5) VALUE SPACES.
01 WS-NUM PIC 9(3) VALUE ZEROS.
PROCEDURE DIVISION.

    IF WS-NAME = SPACES
        DISPLAY "SPACE"
    END-IF.
    IF WS-NUM = ZEROS
        DISPLAY "ZERO"
    END-IF.
    STOP RUN.

