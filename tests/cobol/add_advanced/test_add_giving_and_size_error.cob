*> vybe-test: cobol/add_advanced/test_add_giving_and_size_error
*> origin: languages/cobol/tests/cobol/test_add_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) VALUE 99999.
01 WS-B PIC 9(5) VALUE 0.
PROCEDURE DIVISION.

    ADD WS-A TO WS-A GIVING WS-B ON SIZE ERROR
        DISPLAY WS-B
    END-ADD.
    STOP RUN.

