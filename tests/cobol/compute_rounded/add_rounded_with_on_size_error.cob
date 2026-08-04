*> vybe-test: cobol/compute_rounded/add_rounded_with_on_size_error
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3)V9 VALUE 999.9.
PROCEDURE DIVISION.
    ADD 0.5 TO A ROUNDED
    ON SIZE ERROR
        DISPLAY "OVERFLOW"
    END-ADD.
    STOP RUN.

