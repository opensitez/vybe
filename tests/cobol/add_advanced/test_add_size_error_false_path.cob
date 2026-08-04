*> vybe-test: cobol/add_advanced/test_add_size_error_false_path
*> origin: languages/cobol/tests/cobol/test_add_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 10.
01 WS-B PIC 99 VALUE 0.
PROCEDURE DIVISION.

    ADD WS-A TO WS-B ON SIZE ERROR
        DISPLAY "ERR"
    NOT ON SIZE ERROR
        DISPLAY WS-B
    END-ADD.
    STOP RUN.

