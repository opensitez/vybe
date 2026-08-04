*> vybe-test: cobol/compute_advanced/test_compute_with_end_compute_block
*> origin: languages/cobol/tests/cobol/test_compute_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9 VALUE 3.
01 WS-B PIC 9 VALUE 4.
01 WS-R PIC 9 VALUE 0.
PROCEDURE DIVISION.

    COMPUTE WS-R = WS-A + WS-B
        ON SIZE ERROR
            DISPLAY "ERR"
    END-COMPUTE.
    DISPLAY WS-R.
    STOP RUN.

