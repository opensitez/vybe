*> vybe-test: cobol/compute_rounded/subtract_rounded_with_size_error
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    SUBTRACT 500 FROM A ROUNDED
    ON SIZE ERROR
        DISPLAY "UNDERFLOW"
    END-SUBTRACT.
    STOP RUN.

