*> vybe-test: cobol/compute_advanced/test_compute_size_error_path
*> origin: languages/cobol/tests/cobol/test_compute_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-R PIC 9 VALUE 0.
01 WS-OK PIC X VALUE "N".
PROCEDURE DIVISION.

    COMPUTE WS-R = 999 * 999
        ON SIZE ERROR MOVE "Y" TO WS-OK
        NOT ON SIZE ERROR MOVE "N" TO WS-OK
    END-COMPUTE
    DISPLAY WS-OK.
    STOP RUN.

